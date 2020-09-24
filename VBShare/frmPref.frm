VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPref 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   2955
   ClientLeft      =   5580
   ClientTop       =   3870
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPref.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNorm 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2880
      TabIndex        =   18
      Text            =   "#000000"
      Top             =   900
      Width           =   795
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2040
      TabIndex        =   16
      Text            =   "#000080"
      Top             =   900
      Width           =   795
   End
   Begin VB.TextBox txtCom 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2880
      TabIndex        =   14
      Text            =   "#008000"
      Top             =   480
      Width           =   795
   End
   Begin VB.TextBox txtBG 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2040
      TabIndex        =   12
      Text            =   "#FFFFFF"
      Top             =   480
      Width           =   795
   End
   Begin VB.TextBox txtSig 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   60
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1680
      Width           =   4395
   End
   Begin VB.TextBox RemFoc 
      Height          =   285
      Left            =   4860
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3060
      Width           =   180
   End
   Begin VB.TextBox txtPrev 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmPref.frx":030A
      Top             =   240
      Width           =   1215
   End
   Begin VBShare.lvButtons_H cmdOK 
      Height          =   315
      Left            =   3780
      TabIndex        =   0
      Top             =   60
      Width           =   675
      _ExtentX        =   1191
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
   Begin VBShare.lvButtons_H cmdCancel 
      Height          =   315
      Left            =   3780
      TabIndex        =   1
      Top             =   480
      Width           =   675
      _ExtentX        =   1191
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
   Begin VBShare.lvButtons_H cmdApply 
      Height          =   315
      Left            =   3780
      TabIndex        =   2
      Top             =   900
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "&Apply"
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
   Begin VBShare.lvButtons_H cmdFont 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      Caption         =   "Font"
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
   Begin VBShare.lvButtons_H cmdColor 
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      Caption         =   "Color"
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
   Begin MSComDlg.CommonDialog CD 
      Left            =   4440
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optSig 
      Appearance      =   0  'Flat
      Caption         =   "Bottom"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2700
      TabIndex        =   10
      Top             =   1500
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton optSig 
      Appearance      =   0  'Flat
      Caption         =   "Top"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1860
      TabIndex        =   9
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label lblLoad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "…"
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
      Index           =   3
      Left            =   3480
      TabIndex        =   25
      Top             =   720
      Width           =   195
   End
   Begin VB.Label lblLoad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "…"
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
      Index           =   2
      Left            =   2640
      TabIndex        =   24
      Top             =   720
      Width           =   195
   End
   Begin VB.Label lblLoad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "…"
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
      Index           =   0
      Left            =   2640
      TabIndex        =   23
      Top             =   300
      Width           =   195
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "…"
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
      Left            =   3540
      TabIndex        =   22
      Top             =   300
      Width           =   195
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Placement:"
      Height          =   195
      Index           =   7
      Left            =   900
      TabIndex        =   21
      Top             =   1500
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HTMLColors"
      Height          =   195
      Index           =   6
      Left            =   2040
      TabIndex        =   20
      Top             =   60
      Width           =   1635
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   4440
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   1980
      X2              =   1980
      Y1              =   0
      Y2              =   1260
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      Height          =   195
      Index           =   5
      Left            =   2940
      TabIndex        =   19
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword"
      Height          =   195
      Index           =   4
      Left            =   2100
      TabIndex        =   17
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   195
      Index           =   3
      Left            =   2940
      TabIndex        =   15
      Top             =   300
      Width           =   615
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Code BG"
      Height          =   195
      Index           =   2
      Left            =   2100
      TabIndex        =   13
      Top             =   300
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Author 'Water-Mark'"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   1320
      Width           =   4395
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   3720
      X2              =   3720
      Y1              =   0
      Y2              =   1260
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program Font"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   1875
   End
End
Attribute VB_Name = "frmPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Prefs True
    ApplyFont
End Sub

Private Sub cmdCancel_Click()
    Prefs False
    Unload Me
End Sub

Private Sub cmdColor_Click()
    With CD
        .Color = txtPrev.ForeColor
        .ShowColor
        txtPrev.ForeColor = .Color
    End With
End Sub

Private Sub cmdFont_Click()
    With CD
        .Flags = 1
        .FontName = txtPrev.FontName
        .FontSize = txtPrev.FontSize
        .ShowFont
        txtPrev.FontName = .FontName
        txtPrev.FontSize = Round(.FontSize, 0)
        txtPrev = Round(.FontSize, 0) & " Point" & vbCrLf & .FontName
        txtPrev.FontBold = False
        txtPrev.FontItalic = False
        txtPrev.FontStrikethru = False
        txtPrev.FontUnderline = False
    End With
End Sub

Private Sub cmdOK_Click()
    Prefs True
    ApplyFont
    Unload Me
End Sub

Private Sub Form_Load()
    Prefs False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FullExit = False Then
        Cancel = 1
        Me.Visible = False
    End If
End Sub

Private Sub lblLoad_Click(Index As Integer)
    With CD
        Select Case Index
            Case 0
                .Color = Val("&H" & Mid(txtBG.Text, 2)) * IIf(Left(Val("&H" & Mid(txtBG.Text, 2)), 1) = "-", -1, 1)
            Case 1
                .Color = Val("&H" & Mid(txtCom.Text, 2)) * IIf(Left(Val("&H" & Mid(txtCom.Text, 2)), 1) = "-", -1, 1)
            Case 2
                .Color = Val("&H" & Mid(txtKey.Text, 2)) * IIf(Left(Val("&H" & Mid(txtKey.Text, 2)), 1) = "-", -1, 1)
            Case 3
                .Color = Val("&H" & Mid(txtNorm.Text, 2)) * IIf(Left(Val("&H" & Mid(txtNorm.Text, 2)), 1) = "-", -1, 1)
        End Select
        .ShowColor
        Select Case Index
            Case 0
                txtBG = "#" & LoopString("0", 6 - Len(Hex(.Color))) & Hex(.Color)
            Case 1
                txtCom = "#" & LoopString("0", 6 - Len(Hex(.Color))) & Hex(.Color)
            Case 2
                txtKey = "#" & LoopString("0", 6 - Len(Hex(.Color))) & Hex(.Color)
            Case 3
                txtNorm = "#" & LoopString("0", 6 - Len(Hex(.Color))) & Hex(.Color)
        End Select
    End With
End Sub

Private Sub optSig_Click(Index As Integer)
    If Me.Visible = True Then RemFoc.SetFocus
End Sub

Sub Prefs(Save As Boolean)
    If Save = True Then
        SaveSetting "VBShare", "General", "FontColor", txtPrev.ForeColor
        SaveSetting "VBShare", "General", "FontName", txtPrev.FontName
        SaveSetting "VBShare", "General", "FontSize", Round(txtPrev.FontSize, 0)
        SaveSetting "VBShare", "General", "MarkPos", IIf(optSig(1).Value = True, 1, 0)
        SaveSetting "VBShare", "General", "MarkText", txtSig.Text
        SaveSetting "VBShare", "General", "CodeBG", txtBG.Text
        SaveSetting "VBShare", "General", "Comments", txtCom.Text
        SaveSetting "VBShare", "General", "Keywords", txtKey.Text
        SaveSetting "VBShare", "General", "Normal", txtNorm.Text
    Else
        txtPrev.ForeColor = GetSetting("VBShare", "General", "FontColor", vbWindowText)
        txtPrev.FontName = GetSetting("VBShare", "General", "FontName", "Tahoma")
        txtPrev.FontSize = Round(GetSetting("VBShare", "General", "FontSize", 7), 0)
        txtPrev = Round(txtPrev.FontSize, 0) & " Point" & vbCrLf & txtPrev.FontName
        optSig(Val(GetSetting("VBShare", "General", "MarkPos", 1))).Value = True
        txtSig = GetSetting("VBShare", "General", "MarkText", "")
        txtPrev.FontBold = False
        txtPrev.FontItalic = False
        txtPrev.FontStrikethru = False
        txtPrev.FontUnderline = False
        txtBG = GetSetting("VBShare", "General", "CodeBG", "#FFFFFF")
        txtCom = GetSetting("VBShare", "General", "Comments", "#008000")
        txtKey = GetSetting("VBShare", "General", "Keywords", "#000080")
        txtNorm = GetSetting("VBShare", "General", "Normal", "#000000")
        ApplyFont
    End If
End Sub

Public Sub ApplyFont()
    Dim i As Long
    On Error Resume Next
    txtPrev.FontBold = False
    txtPrev.FontItalic = False
    txtPrev.FontStrikethru = False
    txtPrev.FontUnderline = False
    With frmCode
        For i = 0 To .Count - 1
            .Controls(i).ForeColor = txtPrev.ForeColor
            .Controls(i).Font = txtPrev.Font
        Next i
    End With
    With frmHTML
        For i = 0 To .Count - 1
            .Controls(i).ForeColor = txtPrev.ForeColor
            .Controls(i).Font = txtPrev.Font
        Next i
    End With
    With frmPref
        For i = 0 To .Count - 1
            .Controls(i).ForeColor = txtPrev.ForeColor
            .Controls(i).Font = txtPrev.Font
        Next i
        For i = 0 To 3
            lblLoad(i).FontBold = True
        Next i
    End With
    With frmCode
        For i = 0 To .Count - 1
            .Controls(i).ForeColor = txtPrev.ForeColor
            .Controls(i).Font = txtPrev.Font
        Next i
    End With
    With frmTitle
        For i = 0 To .Count - 1
            .Controls(i).ForeColor = txtPrev.ForeColor
            .Controls(i).Font = txtPrev.Font
        Next i
    End With
    With frmPrint
        For i = 0 To .Count - 1
            .Controls(i).ForeColor = txtPrev.ForeColor
            .Controls(i).Font = txtPrev.Font
        Next i
    End With
End Sub

Private Sub txtBG_LostFocus()
If Left(txtBG, 1) <> "#" Then txtBG = "#" & txtBG
End Sub

Private Sub txtCom_LostFocus()
If Left(txtCom, 1) <> "#" Then txtCom = "#" & txtCom
End Sub

Private Sub txtKey_LostFocus()
If Left(txtKey, 1) <> "#" Then txtKey = "#" & txtKey
End Sub

Private Sub txtNorm_LostFocus()
If Left(txtNorm, 1) <> "#" Then txtNorm = "#" & txtNorm
End Sub
