VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "VBShare"
   ClientHeight    =   9525
   ClientLeft      =   2655
   ClientTop       =   1380
   ClientWidth     =   12465
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   60
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save As..."
         Index           =   1
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "Format Code"
         Shortcut        =   ^F
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuShow 
         Caption         =   "Show Code"
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show HTML"
         Index           =   1
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Preview"
         Index           =   2
         Shortcut        =   {F12}
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPref 
         Caption         =   "Preferences"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbt 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PageTitle As String

Private Sub MDIForm_Load()
    frmPrev.Viewer.Navigate "about:blank"
    Prefs False
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    FullExit = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'the form's sizes werent loading properly, so I just turned off the save custom size.
    'Prefs True
    End
End Sub

Sub Prefs(Save As Boolean)
    If Save = False Then
        Me.Left = GetSetting("VBShare", "WinSize", "MainX", 0)
        Me.Top = GetSetting("VBShare", "WinSize", "MainY", 0)
        Me.Height = GetSetting("VBShare", "WinSize", "MainH", Screen.Height - GetTaskbarHeight)
        Me.Width = GetSetting("VBShare", "WinSize", "MainW", Screen.Width)
        Me.WindowState = GetSetting("VBShare", "WinSize", "MainState", vbMaximized)
        With frmCode
            .Left = GetSetting("VBShare", "WinSize", "CodeX", 100)
            .Top = GetSetting("VBShare", "WinSize", "CodeY", 100)
            .Height = GetSetting("VBShare", "WinSize", "CodeH", frmMain.Height - 950)
            .Width = GetSetting("VBShare", "WinSize", "CodeW", (frmMain.Width - 100) / 2 - 165)
        End With
        With frmHTML
            .Left = GetSetting("VBShare", "WinSize", "HTMLX", (frmMain.Width - 100) / 2 + 50)
            .Top = GetSetting("VBShare", "WinSize", "HTMLY", 100)
            .Height = GetSetting("VBShare", "WinSize", "HTMLH", frmMain.Height - 950)
            .Width = GetSetting("VBShare", "WinSize", "HTMLW", (frmMain.Width - 100) / 2 - 165)
        End With
        With frmPrev
            .Left = GetSetting("VBShare", "WinSize", "PrevX", 1000)
            .Top = GetSetting("VBShare", "WinSize", "PrevY", 1000)
            .Height = GetSetting("VBShare", "WinSize", "PrevH", frmMain.Height - 3000)
            .Width = GetSetting("VBShare", "WinSize", "PrevW", frmMain.Width - 2500)
            .WindowState = GetSetting("VBShare", "WinSize", "PrevState", vbNormal)
        End With
        With frmPref
            .Left = GetSetting("VBShare", "WinSize", "PrefX", frmMain.Width * 0.171875)
            .Top = GetSetting("VBShare", "WinSize", "PrefY", frmMain.Height * 0.25)
        End With
    Else
        SaveSetting "VBShare", "WinSize", "MainState", Me.WindowState
        If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
        SaveSetting "VBShare", "WinSize", "MainX", Me.Left
        SaveSetting "VBShare", "WinSize", "MainY", Me.Top
        SaveSetting "VBShare", "WinSize", "MainH", Me.Height
        SaveSetting "VBShare", "WinSize", "MainW", Me.Width
        With frmCode
            SaveSetting "VBShare", "WinSize", "CodeX", .Left
            SaveSetting "VBShare", "WinSize", "CodeY", .Top
            SaveSetting "VBShare", "WinSize", "CodeH", .Height
            SaveSetting "VBShare", "WinSize", "CodeW", .Width
        End With
        With frmHTML
            SaveSetting "VBShare", "WinSize", "HTMLX", .Left
            SaveSetting "VBShare", "WinSize", "HTMLY", .Top
            SaveSetting "VBShare", "WinSize", "HTMLH", .Height
            SaveSetting "VBShare", "WinSize", "HTMLW", .Width
        End With
        With frmPrev
            SaveSetting "VBShare", "WinSize", "PrevState", .WindowState
            If .WindowState = vbMaximized Or .WindowState = vbMinimized Then .WindowState = vbNormal
            SaveSetting "VBShare", "WinSize", "PrevX", .Left
            SaveSetting "VBShare", "WinSize", "PrevY", .Top
            SaveSetting "VBShare", "WinSize", "PrevH", .Height
            SaveSetting "VBShare", "WinSize", "PrevW", .Width
        End With
        With frmPref
            SaveSetting "VBShare", "WinSize", "PrefX", .Left
            SaveSetting "VBShare", "WinSize", "PrefY", .Top
        End With
    End If
End Sub

Private Sub mnuAbt_Click()
    frmAbout.Show
    frmAbout.SetFocus
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFile_Click()
    If frmCode.Visible = False Or frmCode.txtCode = "" Then
        mnuFormat.Enabled = False
    ElseIf frmCode.Visible = True And frmCode.txtCode <> "" Then
        mnuFormat.Enabled = True
    End If
        
    If frmHTML.Visible = False Or frmHTML.txtHTML = "" Then
        mnuSave(0).Enabled = False
        mnuSave(1).Enabled = False
    ElseIf frmHTML.Visible = True And frmHTML.txtHTML <> "" Then
        mnuSave(0).Enabled = True
        mnuSave(1).Enabled = True
    End If
End Sub

Private Sub mnuFormat_Click()
    If frmCode.txtCode = "" Then
        MsgBox "You have no code to format.", vbCritical, "Error"
        Exit Sub
    End If
    frmTitle.Visible = True
    frmTitle.SetFocus
End Sub

Private Sub mnuOpen_Click()
    Dim Ext As String, strTmp As String, CodeStart As Long, HeaderEnd As String
    With CD
        .Filter = "Form Files (*.frm)|*.frm|Module Files (*.bas)|*.bas|Class Files (*.cls)|*.cls|Control Files (*.ctl)|*.ctl|Property Files (*.pag)|*.pag|Text Files (*.txt)|*.txt"
        .FileName = ""
        .DialogTitle = "Load Code File"
        .ShowOpen
        If FileExist(.FileName) = False Then
            If .FileName <> "" Then MsgBox "Specified file cannot be found.", vbCritical, "Error"
            Exit Sub
        End If
        Ext = FileExt(.FileName)
        If Ext <> "bas" And Ext <> "cls" And Ext <> "frm" And Ext <> "ctl" And Ext <> "pag" And Ext <> "txt" Then
            MsgBox "Specified file is not a valid code file.", vbCritical, "Error"
            Exit Sub
        End If
        strTmp = OpenTxt(.FileName)
        If Ext = "bas" Or Ext = "frm" Or Ext = "cls" Or Ext = "pag" Or Ext = "ctl" Then
            If Ext = "bas" Then HeaderEnd = "Attribute VB_Name = """
            If Ext = "frm" Or Ext = "pag" Or Ext = "ctl" Then HeaderEnd = "Attribute VB_Exposed = "
            If Ext = "cls" Then HeaderEnd = "Attribute VB_Ext_KEY = """
            CodeStart = InStr(1, strTmp, HeaderEnd)
                If CodeStart = 0 Then
                    MsgBox "Specified file is not a valid code file.", vbCritical, "Error"
                    Exit Sub
                End If
            CodeStart = InStr(CodeStart + 20, strTmp, vbCrLf) + 2
        ElseIf Ext = "txt" Then
            Dim validate As VbMsgBoxResult
            validate = FormattedMsg("File has the 'txt' extension, this file may not be a valid code file. Are you sure you want to load it?", vbYesNo, "Load Text File")
            If validate = vbNo Then Exit Sub
            CodeStart = 1
        End If
        frmCode.txtCode.Text = Mid(strTmp, CodeStart)
        PageTitle = FileTitle(.FileTitle)
        strTmp = Empty
        frmCode.Show
        frmCode.SetFocus
        frmPrev.Viewer.Navigate "about:blank"
        frmHTML.txtHTML = Empty
        Unload frmPrev
        Unload frmHTML
    End With
End Sub

Private Sub mnuPref_Click()
    frmPref.Show
    frmPref.SetFocus
End Sub

Private Sub mnuPrint_Click()
    frmPrint.Visible = True
    frmPrint.SetFocus
End Sub

Private Sub mnuSave_Click(Index As Integer)
'Currently there's no distinction between save and save as, so you can add that later.
    If frmHTML.txtHTML = "" Then
        MsgBox "You have nothing to save.", vbCritical, "Error"
        Exit Sub
    End If
    With CD
        .FileName = frmMain.GetFileTitle & ".htm"
        .DialogTitle = "Save HTML"
        .DefaultExt = ".htm"
        .Filter = "HTML Files|*.htm;*.html"
        .ShowSave
        If .FileName = "" Then Exit Sub
        If FileExist(.FileName) = True Then
            Dim validate As VbMsgBoxResult
            validate = FormattedMsg("Specified file already exists, would you like to overwrite?", vbYesNo, "Save HTML")
            If validate = vbNo Then Exit Sub
        End If
        SaveTxt .FileName, frmHTML.txtHTML
    End With
End Sub

Private Sub mnuShow_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCode.Show
            frmCode.SetFocus
        Case 1
            frmHTML.Show
            frmHTML.SetFocus
        Case 2
            If frmHTML.txtHTML = "" Then
                MsgBox "You have nothing to save.", vbCritical, "Error"
                Exit Sub
            End If
            SaveTxt Environ("TMP") & "\preview.htm", frmHTML.txtHTML
            frmPrev.Viewer.Navigate Environ("TMP") & "\preview.htm"
            frmPrev.Show
            frmPrev.SetFocus
    End Select
End Sub

Private Sub mnuView_Click()
    If frmHTML.Visible = False Or frmHTML.txtHTML = "" Then
        mnuShow(2).Enabled = False
    ElseIf frmHTML.Visible = True And frmHTML.txtHTML <> "" Then
        mnuShow(2).Enabled = True
    End If
End Sub

Public Function GetFileTitle() As String
    GetFileTitle = PageTitle
End Function
