Attribute VB_Name = "modSerial"
Option Explicit

Function ValidSerial(Serial As String) As Boolean
    'VBS80-0ZMC0-BLS96-07Y2E
    'used in VBShare
    Dim A As Long, L As Long, Block() As String, Value(3) As Long, CheckVal As Long
    If Len(Serial$) <> 23 Then Exit Function
    Block$() = Split(Serial$, Chr$(45))
    If UBound(Block$) <> 3 Then Exit Function
    If Left$(Serial$, 3) <> Chr$(86) & Chr$(66) & Chr$(83) Then Exit Function
    If InStr(1, Serial$, Chr$(65) & Chr$(78) & Chr$(84)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(65) & Chr$(78) & Chr$(78)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(65) & Chr$(78) & Chr$(89)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(66) & Chr$(76) & Chr$(75)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(66) & Chr$(76) & Chr$(84)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(66) & Chr$(76) & Chr$(85)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(76) & Chr$(85) & Chr$(67)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(78) & Chr$(85) & Chr$(66)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(78) & Chr$(89) & Chr$(67)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(80) & Chr$(85) & Chr$(66)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(82) & Chr$(85) & Chr$(66)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(53) & Chr$(80) & Chr$(89)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(65) & Chr$(65)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(66) & Chr$(66)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(67) & Chr$(67)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(76) & Chr$(76)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(78) & Chr$(78)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(80) & Chr$(80)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(82) & Chr$(82)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(86) & Chr$(86)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(88) & Chr$(88)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(90) & Chr$(90)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(81)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(87)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(79)) <> 0 Then Exit Function
    If InStr(1, Serial$, Chr$(85)) <> 0 And InStr(InStr(1, Serial$, Chr$(85)) + 1, Serial$, Chr$(85)) <> 0 Then Exit Function
    If Left$(Block$(1), 1) = Chr$(48) Then If Asc%(Mid$(Block$(3), 3, 1)) - Asc%(Mid$(Block$(1), 3, 1)) <> 12 Then Exit Function
    For L& = 0 To 3
        If Len(Block$(L&)) <> 5 Then Exit Function
        For A& = 1 To Len(Block$(L&))
            Value&(L&) = Value&(L&) + Asc%(Mid$(Block$(L&), A&, 1))
        Next A&
        Value&(L&) = Value&(L&) + Val#(Chr$(38) & Chr$(72) & Block$(L&)) - 8
        If Value&(L&) < 300 Or Value&(L&) > 500 Then Exit Function
        If IsNumeric(Mid$(Block$(L&), 3, 1)) = True Then Exit Function
        If Asc%(Mid$(Block$(L&), 3, 1)) < 75 Or Asc%(Mid$(Block$(L&), 3, 1)) > 90 Then Exit Function
        Value&(L&) = Value&(L&) - 169 - Asc%(Mid$(Block$(L&), 3, 1))
        CheckVal& = CheckVal& + Val#(Chr$(38) & Chr$(72) & Block$(L))
    Next L
    If CheckVal& > 24 Or InStr(1, CheckVal& / 2, Chr$(46)) <> 0 Then Exit Function
    ValidSerial = Chr$(Value&(0) + 9) & Chr$(Value&(1) + 1) & Chr$(Value&(2) - 7) & Chr$(Value&(3) - 3) & Chr$(Asc%(Mid$(Block$(2), 4, 1)) / 2 + Asc%(Right$(Block$(3), 1)) / 2) = Chr(88) & Chr(77) & Chr(80) & Chr(49) & Chr(63)
End Function
