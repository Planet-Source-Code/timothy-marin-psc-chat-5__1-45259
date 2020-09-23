Attribute VB_Name = "Socket"
Function OpenSocket(Wsck As String, FRM As Form) As Integer
    Dim i As Integer
    For i = 1 To FRM.Controls(Wsck).UBound
        If FRM.Controls(Wsck)(i).State = 0 Then
            OpenSocket = i
            Exit Function
        End If
    Next
    i = FRM.Controls(Wsck).UBound + 1
    Load FRM.Controls(Wsck)(i)
    OpenSocket = i
End Function

Sub SendSocket(Wsck As String, Msg As String, Index As Integer, FRM As Form)
    If FRM.Controls(Wsck)(Index).State = 7 Then
        FRM.Controls(Wsck)(Index).SendData Msg & Chr(0)
    End If
End Sub

Sub SendAllSocket(Wsck As String, Msg As String, Exclude As Integer, FRM As Form)
    Dim i As Integer
    For i = 1 To FRM.Controls(Wsck).UBound
        If i <> Exclude And frmMain.IsAuth(i) <> 0 Then SendSocket Wsck, Msg, i, FRM
    Next
End Sub

'Count How many Times a User is online
Function IsOnline(Wsck As String, IP As String, FRM As Form)
    Dim i As Integer
    Dim Count As Long
    For i = 1 To FRM.Controls(Wsck).UBound
        If FRM.Controls(Wsck)(i).State = 7 And FRM.Controls(Wsck)(i).RemoteHostIP = IP Then
            Count = Count + 1
        End If
    Next
    IsOnline = Count
End Function

