Attribute VB_Name = "modCode"
Option Explicit

Public Function toHTML(Code As String, Optional Title As String = "VB Source Example", Optional BackColor As String = "#FFFFFF", Optional CommentColor As String = "#008000", Optional KeywordColor As String = "#000080", Optional NormalColor As String = "#000000") As String
    Dim Line As String, L As Long, tmpCode As String, tmpLine As String, Ins As Long, fnd As Long, ComIns(1) As Long
    BackColor$ = Replace$(BackColor$, """", "")
    CommentColor$ = Replace$(CommentColor$, """", "")
    KeywordColor$ = Replace$(KeywordColor$, """", "")
    NormalColor$ = Replace$(NormalColor$, """", "")
    If Title$ = "" Then Title = "VB Source Example"
    toHTML$ = "<HTML>" & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & Title & "</TITLE>" & vbCrLf & "<META http-equiv=Content-Type content=""text/html; charset=windows-1252"">" & vbCrLf & "<!--" & vbCrLf & "Code formatted using " & App.Title & ", by unsticky." & vbCrLf & "http://www.unsticky.net" & vbCrLf & "--------------------------------------" & vbCrLf & "Code copyright its creator(s) and subject" & vbCrLf & "to all applicable laws and protections." & vbCrLf & "-->" & vbCrLf & "</HEAD>" & vbCrLf & "<BODY bgColor=""" & BackColor$ & """>"
    toHTML$ = toHTML$ & IIf(frmPref.optSig(0).Value = True, vbCrLf & frmPref.txtSig, "") & vbCrLf & "<FONT Color=""" & NormalColor$ & """ Size=3>" & vbCrLf & "<PRE>"
    Code$ = Replace$(Replace$(Replace$(Code$, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
    For L = 0 To UBound(Split(Code$, vbCrLf))
        Line$ = " " & Split(Code$, vbCrLf)(L) & " "
        ComIns&(0) = InStr(1, Line$, " '")
        If ComIns&(0) <> 0 Then
            If OpenQuotes(Left$(Line$, ComIns&(0) - 1)) = False Then Line$ = Left$(Line$, ComIns&(0)) & "<FONT Color=""" & CommentColor$ & """>" & RTrim$(Mid$(Line$, ComIns&(0) + 1)) & "</FONT>"
        End If
        ComIns&(1) = InStr(1, Line$, " Rem ")
        If ComIns&(1) <> 0 Then
            If OpenQuotes(Left$(Line$, ComIns&(1) - 1)) = False Then Line$ = Left$(Line$, ComIns&(1)) & "<FONT Color=""" & CommentColor$ & """>" & RTrim$(Mid$(Line$, ComIns&(1) + 1)) & "</FONT>"
        End If
        fnd& = InStr(1, Line$, "Optional ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Optional</FONT> " & Mid$(Line$, fnd& + 9)
            fnd& = InStr(fnd& + 1, Line$, "Optional ")
        Loop
        fnd& = InStr(1, Line$, "ByVal ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>ByVal</FONT> " & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, "ByVal ")
        Loop
        fnd& = InStr(1, Line$, "ByRef ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>ByRef</FONT> " & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, "ByRef ")
        Loop
        fnd& = InStr(1, Line$, " As Currency")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Currency</FONT>" & Mid$(Line$, fnd& + 12)
            fnd& = InStr(fnd& + 1, Line$, " As Currency")
        Loop
        fnd& = InStr(1, Line$, " As Boolean")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Boolean</FONT>" & Mid$(Line$, fnd& + 11)
            fnd& = InStr(fnd& + 1, Line$, " As Boolean")
        Loop
        fnd& = InStr(1, Line$, " As Integer")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Integer</FONT>" & Mid$(Line$, fnd& + 11)
            fnd& = InStr(fnd& + 1, Line$, " As Integer")
        Loop
        fnd& = InStr(1, Line$, " As Variant")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Variant</FONT>" & Mid$(Line$, fnd& + 11)
            fnd& = InStr(fnd& + 1, Line$, " As Variant")
        Loop
        fnd& = InStr(1, Line$, " As Double")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Double</FONT>" & Mid$(Line$, fnd& + 10)
            fnd& = InStr(fnd& + 1, Line$, " As Double")
        Loop
        fnd& = InStr(1, Line$, " As Object")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Object</FONT>" & Mid$(Line$, fnd& + 10)
            fnd& = InStr(fnd& + 1, Line$, " As Object")
        Loop
        fnd& = InStr(1, Line$, " As Single")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Single</FONT>" & Mid$(Line$, fnd& + 10)
            fnd& = InStr(fnd& + 1, Line$, " As Single")
        Loop
        fnd& = InStr(1, Line$, " As String")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As String</FONT>" & Mid$(Line$, fnd& + 10)
            fnd& = InStr(fnd& + 1, Line$, " As String")
        Loop
        fnd& = InStr(1, Line$, " As Byte")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Byte</FONT>" & Mid$(Line$, fnd& + 8)
            fnd& = InStr(fnd& + 1, Line$, " As Byte")
        Loop
        fnd& = InStr(1, Line$, " As Date")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Date</FONT>" & Mid$(Line$, fnd& + 8)
            fnd& = InStr(fnd& + 1, Line$, " As Date")
        Loop
        fnd& = InStr(1, Line$, " As Long")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As Long</FONT>" & Mid$(Line$, fnd& + 8)
            fnd& = InStr(fnd& + 1, Line$, " As Long")
        Loop
        fnd& = InStr(1, Line$, " As New ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As New</FONT> " & Mid$(Line$, fnd& + 8)
            fnd& = InStr(fnd& + 1, Line$, " As New ")
        Loop
        fnd& = InStr(1, Line$, " As ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>As</FONT> " & Mid$(Line$, fnd& + 4)
            fnd& = InStr(fnd& + 1, Line$, " As ")
        Loop
        fnd& = InStr(1, Line$, " CCur(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CCur</FONT>(" & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " CCur(")
        Loop
        fnd& = InStr(1, Line$, " CDbl(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CDbl</FONT>(" & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " CDbl(")
        Loop
        fnd& = InStr(1, Line$, " CInt(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CInt</FONT>(" & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " CInt(")
        Loop
        fnd& = InStr(1, Line$, " CSng(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CSng</FONT>(" & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " CSng(")
        Loop
        fnd& = InStr(1, Line$, " CStr(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CStr</FONT>(" & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " CStr(")
        Loop
        fnd& = InStr(1, Line$, " CVar(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CVar</FONT>(" & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " CVar(")
        Loop
        fnd& = InStr(1, Line$, " CDec(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CDec</FONT>(" & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " CDec(")
        Loop
        fnd& = InStr(1, Line$, " CBool(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CBool</FONT>(" & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " CBool(")
        Loop
        fnd& = InStr(1, Line$, " CDate(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CDate</FONT>(" & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " CDate(")
        Loop
        fnd& = InStr(1, Line$, " CByte(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>CByte</FONT>(" & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " CByte(")
        Loop
        fnd& = InStr(1, Line$, " Spc(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Spc</FONT>(" & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Spc(")
        Loop
        fnd& = InStr(1, Line$, " Tab(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Tab</FONT>(" & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Tab(")
        Loop
        fnd& = InStr(1, Line$, "Exit Function ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Exit Function</FONT> " & Mid$(Line$, fnd& + 14)
            fnd& = InStr(fnd& + 1, Line$, "Exit Function ")
        Loop
        fnd& = InStr(1, Line$, "Exit Property ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Exit Property</FONT> " & Mid$(Line$, fnd& + 14)
            fnd& = InStr(fnd& + 1, Line$, "Exit Property ")
        Loop
        fnd& = InStr(1, Line$, "Exit Select ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Exit Select</FONT> " & Mid$(Line$, fnd& + 12)
            fnd& = InStr(fnd& + 1, Line$, "Exit Select ")
        Loop
        fnd& = InStr(1, Line$, "Exit Next ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Exit Next</FONT> " & Mid$(Line$, fnd& + 10)
            fnd& = InStr(fnd& + 1, Line$, "Exit Next ")
        Loop
        fnd& = InStr(1, Line$, "Exit With ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Exit With</FONT> " & Mid$(Line$, fnd& + 10)
            fnd& = InStr(fnd& + 1, Line$, "Exit With ")
        Loop
        fnd& = InStr(1, Line$, "Exit Sub ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Exit Sub</FONT> " & Mid$(Line$, fnd& + 9)
            fnd& = InStr(fnd& + 1, Line$, "Exit Sub ")
        Loop
        fnd& = InStr(1, Line$, "Exit For ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Exit For</FONT> " & Mid$(Line$, fnd& + 9)
            fnd& = InStr(fnd& + 1, Line$, "Exit For ")
        Loop
            fnd& = InStr(1, Line$, "Exit Do ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Exit Do</FONT> " & Mid$(Line$, fnd& + 8)
            fnd& = InStr(fnd& + 1, Line$, "Exit Do ")
        Loop
        fnd& = InStr(1, Line$, " For Append ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>For Append</FONT> " & Mid$(Line$, fnd& + 12)
            fnd& = InStr(fnd& + 1, Line$, " For Append ")
        Loop
        fnd& = InStr(1, Line$, " For Random ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>For Random</FONT> " & Mid$(Line$, fnd& + 12)
            fnd& = InStr(fnd& + 1, Line$, " For Random ")
        Loop
        fnd& = InStr(1, Line$, " For Output ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>For Output</FONT> " & Mid$(Line$, fnd& + 12)
            fnd& = InStr(fnd& + 1, Line$, " For Output ")
        Loop
        fnd& = InStr(1, Line$, " For Binary ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>For Binary</FONT> " & Mid$(Line$, fnd& + 12)
            fnd& = InStr(fnd& + 1, Line$, " For Binary ")
        Loop
        fnd& = InStr(1, Line$, " For Input ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>For Input</FONT> " & Mid$(Line$, fnd& + 11)
            fnd& = InStr(fnd& + 1, Line$, " For Input ")
        Loop
        fnd& = InStr(1, Line$, " For Each ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>For Each</FONT> " & Mid$(Line$, fnd& + 10)
            fnd& = InStr(fnd& + 1, Line$, " For Each ")
        Loop
        fnd& = InStr(1, Line$, " For ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>For</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " For ")
        Loop
        fnd& = InStr(1, Line$, " Line Input ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Line Input</FONT> " & Mid$(Line$, fnd& + 12)
            fnd& = InStr(fnd& + 1, Line$, " Line Input ")
        Loop
        fnd& = InStr(1, Line$, " Input ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Input</FONT> " & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " Input ")
        Loop
        fnd& = InStr(1, Line$, " Print ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Print</FONT> " & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " Print ")
        Loop
        fnd& = InStr(1, Line$, ".Print ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & ".<FONT Color=""" & KeywordColor$ & """>Print</FONT> " & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " Print ")
        Loop
        fnd& = InStr(1, Line$, " Write ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Write</FONT> " & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " Write ")
        Loop
        fnd& = InStr(1, Line$, " Raise Event ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Raise Event</FONT> " & Mid$(Line$, fnd& + 13)
            fnd& = InStr(fnd& + 1, Line$, " Raise Event ")
        Loop
        fnd& = InStr(1, Line$, " Debug.Print ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Debug.Print</FONT> " & Mid$(Line$, fnd& + 13)
            fnd& = InStr(fnd& + 1, Line$, " Debug.Print ")
        Loop
        fnd& = InStr(1, Line$, " Then ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Then</FONT> " & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " Then ")
        Loop
        fnd& = InStr(1, Line$, " True ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>True</FONT> " & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " True ")
        Loop
        fnd& = InStr(1, Line$, " False ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>False</FONT> " & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " False ")
        Loop
        fnd& = InStr(1, Line$, " And ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>And</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " And ")
        Loop
        fnd& = InStr(1, Line$, " Not ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Not</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Not ")
        Loop
        fnd& = InStr(1, Line$, " Xor ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Xor</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Xor ")
        Loop
        fnd& = InStr(1, Line$, " Or ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Or</FONT> " & Mid$(Line$, fnd& + 4)
            fnd& = InStr(fnd& + 1, Line$, " Or ")
        Loop
        fnd& = InStr(1, Line$, " To ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>To</FONT> " & Mid$(Line$, fnd& + 4)
            fnd& = InStr(fnd& + 1, Line$, " To ")
        Loop
        fnd& = InStr(1, Line$, " In ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>In</FONT> " & Mid$(Line$, fnd& + 4)
            fnd& = InStr(fnd& + 1, Line$, " In ")
        Loop
        fnd& = InStr(1, Line$, " Is ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Is</FONT> " & Mid$(Line$, fnd& + 4)
            fnd& = InStr(fnd& + 1, Line$, " Is ")
        Loop
        fnd& = InStr(1, Line$, " Like ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Like</FONT> " & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " Like ")
        Loop
        fnd& = InStr(1, Line$, " Eqv ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Eqv</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Eqv ")
        Loop
        fnd& = InStr(1, Line$, " Imp ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Imp</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Imp ")
        Loop
        fnd& = InStr(1, Line$, " Mod ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Mod</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Mod ")
        Loop
        fnd& = InStr(1, Line$, " LBound(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>LBound</FONT>(" & Mid$(Line$, fnd& + 8)
            fnd& = InStr(fnd& + 1, Line$, " LBound(")
        Loop
        fnd& = InStr(1, Line$, " UBound(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>UBound</FONT>(" & Mid$(Line$, fnd& + 8)
            fnd& = InStr(fnd& + 1, Line$, " UBound(")
        Loop
        fnd& = InStr(1, Line$, " StrComp(")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """> StrComp</FONT>(" & Mid$(Line$, fnd& + 9)
            fnd& = InStr(fnd& + 1, Line$, " StrComp(")
        Loop
        fnd& = InStr(1, Line$, " Call ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Call</FONT> " & Mid$(Line$, fnd& + 6)
            fnd& = InStr(fnd& + 1, Line$, " Call ")
        Loop
        fnd& = InStr(1, Line$, " Set ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Set</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Set ")
        Loop
        fnd& = InStr(1, Line$, " Let ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Let</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Let ")
        Loop
        fnd& = InStr(1, Line$, " Alias ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Alias</FONT> " & Mid$(Line$, fnd& + 7)
            fnd& = InStr(fnd& + 1, Line$, " Alias ")
        Loop
        fnd& = InStr(1, Line$, " Lib ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>Lib</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " Lib ")
        Loop
        fnd& = InStr(1, Line$, "Option Explicit ")
            If Left$(LTrim$(Line), 16) = "Option Explicit " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Option Explicit</FONT> " & Mid$(Line$, fnd& + 16)
        fnd& = InStr(1, Line$, "Option Compare ")
            If Left$(LTrim$(Line), 15) = "Option Compare " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Option Compare</FONT> " & Mid$(Line$, fnd& + 15)
        fnd& = InStr(1, Line$, "Option Private ")
            If Left$(LTrim$(Line), 15) = "Option Private " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Option Private</FONT> " & Mid$(Line$, fnd& + 15)
        fnd& = InStr(1, Line$, "Option Base ")
            If Left$(LTrim$(Line), 12) = "Option Base " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Option Base</FONT> " & Mid$(Line$, fnd& + 12)
        fnd& = InStr(1, Line$, "Private Declare Function ")
            If Left$(LTrim$(Line), 25) = "Private Declare Function " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Private Declare Function</FONT> " & Mid$(Line$, fnd& + 25)
        fnd& = InStr(1, Line$, "Private Declare Sub ")
            If Left$(LTrim$(Line), 20) = "Private Declare Sub " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Private Declare Sub</FONT> " & Mid$(Line$, fnd& + 20)
        fnd& = InStr(1, Line$, "Private Function ")
            If Left$(LTrim$(Line), 17) = "Private Function " Then Line$ = Left$(Line$, fnd& - 1) & "<HR Color=""" & NormalColor$ & """ Size=1>" & vbCrLf & "<FONT Color=""" & KeywordColor$ & """>Private Function</FONT> " & Mid$(Line$, fnd& + 17)
        fnd& = InStr(1, Line$, "Private Property ")
            If Left$(LTrim$(Line), 17) = "Private Property " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Private Property</FONT> " & Mid$(Line$, fnd& + 17)
        fnd& = InStr(1, Line$, "Private Const ")
            If Left$(LTrim$(Line), 14) = "Private Const " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Private Const</FONT> " & Mid$(Line$, fnd& + 14)
        fnd& = InStr(1, Line$, "Private Event ")
            If Left$(LTrim$(Line), 14) = "Private Event " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Private Event</FONT> " & Mid$(Line$, fnd& + 14)
        fnd& = InStr(1, Line$, "Private Type ")
            If Left$(LTrim$(Line), 13) = "Private Type " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Private Type</FONT> " & Mid$(Line$, fnd& + 13)
        fnd& = InStr(1, Line$, "Private Sub ")
            If Left$(LTrim$(Line), 12) = "Private Sub " Then Line$ = Left$(Line$, fnd& - 1) & "<HR Color=""" & NormalColor$ & """ Size=1>" & vbCrLf & "<FONT Color=""" & KeywordColor$ & """>Private Sub</FONT> " & Mid$(Line$, fnd& + 12)
        fnd& = InStr(1, Line$, "Public Declare Function ")
            If Left$(LTrim$(Line), 24) = "Public Declare Function " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Public Declare Function</FONT> " & Mid$(Line$, fnd& + 24)
            fnd& = InStr(1, Line$, "Public Declare Sub ")
            If Left$(LTrim$(Line), 19) = "Public Declare Sub " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Public Declare Sub</FONT> " & Mid$(Line$, fnd& + 19)
        fnd& = InStr(1, Line$, "Public Function ")
            If Left$(LTrim$(Line), 16) = "Public Function " Then Line$ = Left$(Line$, fnd& - 1) & "<HR Color=""" & NormalColor$ & """ Size=1>" & vbCrLf & "<FONT Color=""" & KeywordColor$ & """>Public Function</FONT> " & Mid$(Line$, fnd& + 16)
        fnd& = InStr(1, Line$, "Public Property ")
            If Left$(LTrim$(Line), 16) = "Public Property " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Public Property</FONT> " & Mid$(Line$, fnd& + 16)
        fnd& = InStr(1, Line$, "Public Const ")
            If Left$(LTrim$(Line), 13) = "Public Const " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Public Const</FONT> " & Mid$(Line$, fnd& + 13)
        fnd& = InStr(1, Line$, "Public Event ")
            If Left$(LTrim$(Line), 13) = "Public Event " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Public Event</FONT> " & Mid$(Line$, fnd& + 13)
        fnd& = InStr(1, Line$, "Public Type ")
            If Left$(LTrim$(Line), 12) = "Public Type " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Public Type</FONT> " & Mid$(Line$, fnd& + 12)
        fnd& = InStr(1, Line$, "Public Sub ")
            If Left$(LTrim$(Line), 11) = "Public Sub " Then Line$ = Left$(Line$, fnd& - 1) & "<HR Color=""" & NormalColor$ & """ Size=1>" & vbCrLf & "<FONT Color=""" & KeywordColor$ & """>Public Sub</FONT> " & Mid$(Line$, fnd& + 11)
        fnd& = InStr(1, Line$, "Declare Function ")
            If Left$(LTrim$(Line), 17) = "Declare Function " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Declare Function</FONT> " & Mid$(Line$, fnd& + 17)
        fnd& = InStr(1, Line$, "Declare Sub ")
            If Left$(LTrim$(Line), 12) = "Declare Sub " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Declare Sub</FONT> " & Mid$(Line$, fnd& + 12)
        fnd& = InStr(1, Line$, "Function ")
            If Left$(LTrim$(Line), 9) = "Function " Then Line$ = Left$(Line$, fnd& - 1) & "<HR Color=""" & NormalColor$ & """ Size=1>" & vbCrLf & "<FONT Color=""" & KeywordColor$ & """>Function</FONT> " & Mid$(Line$, fnd& + 9)
        fnd& = InStr(1, Line$, "Property ")
            If Left$(LTrim$(Line), 9) = "Property " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Property</FONT> " & Mid$(Line$, fnd& + 9)
        fnd& = InStr(1, Line$, "Const ")
            If Left$(LTrim$(Line), 6) = "Const " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Const</FONT> " & Mid$(Line$, fnd& + 6)
        fnd& = InStr(1, Line$, "#Const ")
            If Left$(LTrim$(Line), 7) = "#Const " Then Line$ = Left$(Line$, fnd& - 1) & "#<FONT Color=" & KeywordColor$ & """>Const</FONT> " & Mid$(Line$, fnd& + 7)
        fnd& = InStr(1, Line$, "Event ")
            If Left$(LTrim$(Line), 6) = "Event " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Event</FONT> " & Mid$(Line$, fnd& + 6)
        fnd& = InStr(1, Line$, "Type ")
            If Left$(LTrim$(Line), 5) = "Type " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Type</FONT> " & Mid$(Line$, fnd& + 5)
        fnd& = InStr(1, Line$, "Sub ")
            If Left$(LTrim$(Line), 4) = "Sub " Then Line$ = Left$(Line$, fnd& - 1) & "<HR Color=""" & NormalColor$ & """ Size=1>" & vbCrLf & "<FONT Color=""" & KeywordColor$ & """>Sub</FONT> " & Mid$(Line$, fnd& + 4)
        fnd& = InStr(1, Line$, "Private ")
            If Left$(LTrim$(Line), 8) = "Private " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Private</FONT> " & Mid$(Line$, fnd& + 8)
        fnd& = InStr(1, Line$, "Public ")
            If Left$(LTrim$(Line), 7) = "Public " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Public</FONT> " & Mid$(Line$, fnd& + 7)
        fnd& = InStr(1, Line$, "End Function")
            If Left$(LTrim$(Line), 12) = "End Function" Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>End Function</FONT>" & Mid$(Line$, fnd& + 12)
        fnd& = InStr(1, Line$, "End Property")
            If Left$(LTrim$(Line), 12) = "End Property" Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>End Property</FONT>" & Mid$(Line$, fnd& + 12)
        fnd& = InStr(1, Line$, "End Select")
            If Left$(LTrim$(Line), 10) = "End Select" Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>End Select</FONT>" & Mid$(Line$, fnd& + 10)
        fnd& = InStr(1, Line$, "End Type")
            If Left$(LTrim$(Line), 8) = "End Type" Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>End Type</FONT>" & Mid$(Line$, fnd& + 8)
        fnd& = InStr(1, Line$, "End With")
            If Left$(LTrim$(Line), 8) = "End With" Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>End With</FONT>" & Mid$(Line$, fnd& + 8)
        fnd& = InStr(1, Line$, "End Sub")
            If Left$(LTrim$(Line), 7) = "End Sub" Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>End Sub</FONT>" & Mid$(Line$, fnd& + 7)
        fnd& = InStr(1, Line$, "#End If")
            If Left$(LTrim$(Line), 7) = "#End If" Then Line$ = Left$(Line$, fnd& - 1) & "#<FONT Color=" & KeywordColor$ & """>End If</FONT>" & Mid$(Line$, fnd& + 7)
        fnd& = InStr(1, Line$, "End If")
            If Left$(LTrim$(Line), 6) = "End If" Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>End If</FONT>" & Mid$(Line$, fnd& + 6)
        fnd& = InStr(1, Line$, "On Error Resume Next")
            If Left$(LTrim$(Line), 20) = "On Error Resume Next" Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>On Error Resume Next</FONT>" & Mid$(Line$, fnd& + 20)
        fnd& = InStr(1, Line$, "On Error Goto ")
            If Left$(LTrim$(Line), 14) = "On Error Goto " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>On Error Goto</FONT> " & Mid$(Line$, fnd& + 14)
        fnd& = InStr(1, Line$, "#ElseIf ")
            If Left$(LTrim$(Line), 8) = "#ElseIf " Then Line$ = Left$(Line$, fnd& - 1) & "#<FONT Color=" & KeywordColor$ & """>ElseIf</FONT> " & Mid$(Line$, fnd& + 8)
        fnd& = InStr(1, Line$, "ElseIf ")
            If Left$(LTrim$(Line), 7) = "ElseIf " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>ElseIf</FONT> " & Mid$(Line$, fnd& + 7)
        fnd& = InStr(1, Line$, "#Else ")
            If Left$(LTrim$(Line), 6) = "#Else " Then Line$ = Left$(Line$, fnd& - 1) & "#<FONT Color=" & KeywordColor$ & """>Else</FONT> " & Mid$(Line$, fnd& + 6)
        fnd& = InStr(1, Line$, "Else ")
            If Left$(LTrim$(Line), 5) = "Else " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Else</FONT> " & Mid$(Line$, fnd& + 5)
        fnd& = InStr(1, Line$, "Case Else ")
            If Left$(LTrim$(Line), 10) = "Case Else " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Case Else</FONT> " & Mid$(Line$, fnd& + 10)
        fnd& = InStr(1, Line$, "Case ")
            If Left$(LTrim$(Line), 5) = "Case " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Case</FONT> " & Mid$(Line$, fnd& + 5)
        fnd& = InStr(1, Line$, "Do Until ")
            If Left$(LTrim$(Line), 9) = "Do Until " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Do Until</FONT> " & Mid$(Line$, fnd& + 9)
        fnd& = InStr(1, Line$, "Do While")
            If Left$(LTrim$(Line), 9) = "Do While " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Do While</FONT> " & Mid$(Line$, fnd& + 9)
        fnd& = InStr(1, Line$, "Do ")
            If Left$(LTrim$(Line), 6) = "Do " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Do</FONT> " & Mid$(Line$, fnd& + 3)
        fnd& = InStr(1, Line$, "Loop Until ")
            If Left$(LTrim$(Line), 9) = "Loop Until " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Loop Until</FONT> " & Mid$(Line$, fnd& + 11)
        fnd& = InStr(1, Line$, "Loop While ")
            If Left$(LTrim$(Line), 7) = "Loop While " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Loop While</FONT> " & Mid$(Line$, fnd& + 11)
        fnd& = InStr(1, Line$, "Loop ")
            If Left$(LTrim$(Line), 6) = "Loop " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Loop</FONT> " & Mid$(Line$, fnd& + 5)
        fnd& = InStr(1, Line$, "Select Case ")
            If Left$(LTrim$(Line), 12) = "Select Case " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Select Case</FONT> " & Mid$(Line$, fnd& + 12)
        fnd& = InStr(1, Line$, "Randomize ")
            If Left$(LTrim$(Line), 10) = "Randomize " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Randomize</FONT> " & Mid$(Line$, fnd& + 10)
        fnd& = InStr(1, Line$, "DoEvents ")
            If Left$(LTrim$(Line), 9) = "DoEvents " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>DoEvents</FONT> " & Mid$(Line$, fnd& + 9)
        fnd& = InStr(1, Line$, "Global ")
            If Left$(LTrim$(Line), 7) = "Global " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Global</FONT> " & Mid$(Line$, fnd& + 7)
        fnd& = InStr(1, Line$, "Static ")
            If Left$(LTrim$(Line), 7) = "Static " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Static</FONT> " & Mid$(Line$, fnd& + 7)
        fnd& = InStr(1, Line$, "Close ")
            If Left$(LTrim$(Line), 6) = "Close " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Close</FONT> " & Mid$(Line$, fnd& + 6)
        fnd& = InStr(1, Line$, "While ")
            If Left$(LTrim$(Line), 6) = "While " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>While</FONT> " & Mid$(Line$, fnd& + 6)
        fnd& = InStr(1, Line$, "Open ")
            If Left$(LTrim$(Line), 5) = "Open " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Open</FONT> " & Mid$(Line$, fnd& + 5)
        fnd& = InStr(1, Line$, "With ")
            If Left$(LTrim$(Line), 5) = "With " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>With</FONT> " & Mid$(Line$, fnd& + 5)
        fnd& = InStr(1, Line$, "Next ")
            If Left$(LTrim$(Line), 5) = "Next " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Next</FONT> " & Mid$(Line$, fnd& + 5)
        fnd& = InStr(1, Line$, "Wend ")
            If Left$(LTrim$(Line), 5) = "Wend " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Wend</FONT> " & Mid$(Line$, fnd& + 5)
        fnd& = InStr(1, Line$, "Dim ")
            If Left$(LTrim$(Line), 4) = "Dim " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>Dim</FONT> " & Mid$(Line$, fnd& + 4)
        fnd& = InStr(1, Line$, "If ")
            If Left$(LTrim$(Line), 3) = "If " Then Line$ = Left$(Line$, fnd& - 1) & "<FONT Color=""" & KeywordColor$ & """>If</FONT> " & Mid$(Line$, fnd& + 3)
        fnd& = InStr(1, Line$, " End ")
        Do Until fnd& = 0
            If OpenQuotes(Left$(Line$, fnd& - 1)) = False And ValCheck(fnd&, ComIns&(0), ComIns&(1)) = True Then Line$ = Left$(Line$, fnd& - 1) & " <FONT Color=""" & KeywordColor$ & """>End</FONT> " & Mid$(Line$, fnd& + 5)
            fnd& = InStr(fnd& + 1, Line$, " End ")
        Loop
        tmpCode$ = tmpCode$ & vbCrLf & RTrim$(Mid$(Line$, 2))
        DoEvents
    Next L
    toHTML$ = toHTML$ & tmpCode$ & vbCrLf & "</PRE>"
    toHTML$ = toHTML$ & IIf(frmPref.optSig(1).Value = True, vbCrLf & frmPref.txtSig, "") & vbCrLf & "</FONT>" & vbCrLf & "</BODY>" & vbCrLf & "</HTML>"
End Function

Private Function ValCheck(Value As Long, ByVal ToCheck1 As Long, ByVal ToCheck2 As Long) As Boolean
    If ToCheck1 = 0 Then ToCheck1 = Value + 1
    If ToCheck2 = 0 Then ToCheck2 = Value + 1
    ValCheck = Value < ToCheck1
    If ValCheck = False Then Exit Function
    ValCheck = Value < ToCheck2
End Function
