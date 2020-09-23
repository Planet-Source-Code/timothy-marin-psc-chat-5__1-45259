Attribute VB_Name = "Functions"
Public AlertCount As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" _
(lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private m_snd() As Byte
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_MEMORY = &H4 ' lpszSoundName points to a memory file
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound

Public Function PlaySound(ByVal SndID As Long) As Long
      Const Flags = SND_ASYNC Or SND_MEMORY
      m_snd = LoadResData(SndID, "CUSTOM")
      PlaySoundData m_snd(0), 0, Flags
End Function
Public Function DectoWebCol(lngColour As Long) As String
    Dim strColour As String
    'Convert decimal colour to hex
    strColour = Hex(lngColour)
    'Add leading zero's


    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop
    'Reverse the bgr string pairs to rgb
    DectoWebCol = "#" & Right$(strColour, 2) & _
    Mid$(strColour, 3, 2) & _
    Left$(strColour, 2)
End Function

Public Function DisplayAlert(MessageText As String, Duration As Long, Sound As Long)
    Dim AlertBox As frmNotify
    Set AlertBox = New frmNotify
  AlertBox.Display MessageText, Duration, Sound
Exit Function
errtrap:
   MsgBox Err.Description
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
    Msg = Replace(Msg, Chr(60), "")
    Msg = Replace(Msg, Chr(62), "")
    IChr = Msg
End Function
