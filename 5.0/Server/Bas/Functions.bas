Attribute VB_Name = "Functions"
'Generates Random 10 chr Key
Function GRand()
    Randomize
    Dim i As Integer
    Dim z As Integer
    For i = 1 To 10
    z = Int(Rnd * 85) + 45
    GRand = GRand & Chr(z)
    Next
End Function

'Checks Msg For Invalid Chrs allowd(32-127)
Function IChr(Msg As String)
    Dim i As Integer
    For i = 0 To 31
        Msg = Replace(Msg, Chr(i), "")
    Next
    For i = 127 To 255
        Msg = Replace(Msg, Chr(i), "")
    Next
    Msg = Replace(Msg, Chr(60), "&lt;")
    Msg = Replace(Msg, Chr(62), "&gt;")
    IChr = Msg
End Function
