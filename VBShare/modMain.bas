Attribute VB_Name = "modMain"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Global FullExit As Boolean

Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long, rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

Public Function lv_TimerCallBack(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tgtButton As lvButtons_H
    CopyMemory tgtButton, GetProp(hwnd, "lv_ClassID"), &H4
    Call tgtButton.TimerUpdate(GetProp(hwnd, "lv_TimerID"))
    CopyMemory tgtButton, 0&, &H4
End Function

Public Function OpenSite(Site As String)
    On Error Resume Next
    Call ShellExecute(0&, "OPEN", Site, vbNullString, vbNullString, vbNormalFocus)
End Function

Public Function SaveTxt(File As String, Text As String)
    On Error Resume Next
    Kill File
    Open File For Output As #1
        Print #1, Text
    Close #1
End Function

Public Function OpenTxt(File As String) As String
    Dim tmp As String, Line As String
    Open File For Input As #1
    Do Until EOF(1)
        Line Input #1, Line
        tmp = tmp & IIf(tmp <> "", vbCrLf, "") & Line
        DoEvents
    Loop
    Close #1
    OpenTxt = tmp
End Function

Public Function OpenQuotes(toCheck As String) As Boolean
    If toCheck = "" Then Exit Function
    OpenQuotes = InStr(1, UBound(Split(toCheck, """")) / 2, ".") <> 0
End Function

Public Function CountInst(sString As String, sFind As String) As Long
    CountInst = UBound(Split(sString, sFind))
End Function

Public Function FileExist(FPath As String) As Boolean
    If InStr(1, FPath, ":\") = 0 Then FPath = "?"
    FileExist = (Dir(FPath) <> "")
End Function

Public Function FileExt(Path As String) As String
    Dim strSplit() As String
    strSplit() = Split(Path, "\")
    strSplit() = Split(strSplit(UBound(strSplit)), ".")
    FileExt = LCase(strSplit(UBound(strSplit)))
End Function

Public Function FileTitle(Path As String) As String
    Dim strSplit() As String
    strSplit() = Split(Path, "\")
    strSplit() = Split(strSplit(UBound(strSplit)), ".")
    FileTitle = strSplit(UBound(strSplit) - 1)
End Function

Public Function LoopString(strLoop As String, Times As Long) As String
    LoopString = Replace(Space(Times), " ", strLoop)
End Function

Public Function FormattedMsg(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "MsgBox") As VbMsgBoxResult
    If Len(Prompt) > 80 Then
        Dim L As Long, R As Long, LChr As String, RChr As String, Break As Long, tmp(1) As String
        tmp(0) = Prompt
        Prompt = ""
        Do Until tmp(0) = ""
            L = 69
            R = 70
            LChr = Mid(tmp(0), L, 1)
            RChr = Mid(tmp(0), R, 1)
            Do Until LChr = " " Or LChr = "-" Or LChr = "_" Or L = 1
                L = L - 1
                LChr = Mid(tmp(0), L, 1)
            Loop
            Do Until RChr = " " Or RChr = "-" Or RChr = "_" Or R = Len(tmp(0))
                R = R + 1
                RChr = Mid(tmp(0), R, 1)
            Loop
            Break = IIf(70 - L < R - 70, L, R)
            tmp(1) = Left(tmp(0), Break) & vbCrLf
            tmp(1) = IIf(Left(tmp(1), 1) <> " ", tmp(1), Mid(tmp(1), 2))
            tmp(0) = Mid(tmp(0), Break)
            If Len(tmp(0)) < 76 Then
                tmp(1) = tmp(1) & IIf(Left(tmp(0), 1) <> " ", tmp(0), Mid(tmp(0), 2))
                tmp(0) = ""
            End If
            Prompt = Prompt & tmp(1)
        Loop
    End If
    FormattedMsg = MsgBox(Prompt, Buttons, Title)
End Function
