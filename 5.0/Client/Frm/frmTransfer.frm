VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   5040
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar P1 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2445
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   5040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Connect Info"
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      Begin VB.Timer tmrInfo 
         Interval        =   1000
         Left            =   4560
         Top             =   120
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send File"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4080
         TabIndex        =   12
         Top             =   255
         Width           =   1335
      End
      Begin VB.TextBox txtLPort 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Text            =   "111"
         Top             =   675
         Width           =   735
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Listen"
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   675
         Width           =   2895
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Text            =   "111"
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Text            =   "127.0.0.1"
         Top             =   285
         Width           =   1695
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   285
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   3960
         X2              =   3960
         Y1              =   960
         Y2              =   240
      End
      Begin VB.Label lblStatus 
         Caption         =   "Disconnected..."
         Height          =   255
         Left            =   4110
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Stat 
      Caption         =   "Time Remaining :"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   13
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label Stat 
      Caption         =   "Percent :"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Label Stat 
      Caption         =   "Speed :"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label Stat 
      Caption         =   "Size :"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Stat 
      Caption         =   "Filename :"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   5535
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Filename As String
Public FileSize As Long
Public BytesReceived As Long
Public Status As String
Dim FileNum As Integer
Dim GSF As Long
Dim LCHECK As Long

Dim Delay As Boolean

Dim SSF As Long
Dim Chunka As Long

Private Sub cmdConnect_Click()
    On Error GoTo Errors
    Socket.Close
    Socket.Connect txtIP.Text, txtPort.Text
    State "Connecting..."
    Exit Sub
Errors:
    MsgBox err.Description, , "Error"
End Sub

Private Sub cmdListen_Click()
    On Error GoTo Errors
    Socket.Close
    Socket.LocalPort = txtLPort.Text
    Socket.Listen
    State "Listening..."
    Exit Sub
Errors:
    MsgBox err.Description, , "Error"
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
    CD.ShowOpen
    If Len(CD.Filename) > 0 Then
        cmdSend.Enabled = False
        FileNum = FreeFile
        Open CD.Filename For Binary As #FileNum
        FileSize = LOF(FileNum)
        Filename = Right(CD.Filename, Len(CD.Filename) - InStrRev(CD.Filename, "\"))
        State "Requesting..."
        Socket.SendData "SND|" & Filename & "|" & FileSize
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #FileNum
End Sub

Private Sub Socket_Close()
    Close #FileNum
    Socket.Close
    State "Disconnected..."
    cmdConnect.Enabled = True
    cmdListen.Enabled = True
End Sub

Private Sub Socket_Connect()
    State "Connected..."
End Sub

Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
    Socket.Close
    Socket.Accept requestID
    State "Connected..."
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    Dim Sata As String
    Dim Info() As String
    Dim Answer As VbMsgBoxResult
    Socket.GetData Sata
    DoEvents
    Select Case UCase(Left(Sata, 3))
        Case "SND"
            If Status <> "Connected..." Then Exit Sub
            Info = Split(Sata, "|")
            If UBound(Info) < 2 Then
                Socket.SendData "DFT"
                Exit Sub
            End If
            Me.Show
            Answer = MsgBox("Will You Accept " & Info(1) & " : " & DoConv(Info(2)) & " from " & Socket.RemoteHostIP & " ?", vbYesNo, "Tranfer Request")
            If Answer = vbYes Then
                cmdSend.Enabled = False
                State "Waiting..."
                FileNum = FreeFile
                BytesReceived = 0
                LCHECK = 0
                GSF = 0
                FileSize = Int(Info(2))
                Close #FileNum
                Filename = Info(1)
                Open App.Path & "\" & Info(1) For Binary As #FileNum
                If LOF(FileNum) > 0 Then
                    Answer = MsgBox("The file already exists, Would you like to resume?", vbYesNo, "Resume")
                    If Answer = vbYes Then
                        BytesReceived = LOF(FileNum)
                        Socket.SendData "SFT|" & LOF(FileNum)
                    Else
                        Close #FileNum
                        Kill App.Path & "\" & Info(1)
                        DoEvents
                        Open App.Path & "\" & Info(1) For Binary As #FileNum
                        Socket.SendData "SFT|1"
                    End If
                Else
                    Socket.SendData "SFT|1"
                End If
            Else
                Socket.SendData "DFT"
            End If
        Case "SFT"
            If Status <> "Requesting..." Then Exit Sub
            Info = Split(Sata, "|")
            If UBound(Info) < 1 Then
                State "Connected..."
                Exit Sub
            End If
            State "Sending..."
            Dim chunk As String
            BytesReceived = 0
            If LOF(FileNum) < 1024 Then
                chunk = Space$(LOF(FileNum))
            Else
                chunk = Space$(1024)
            End If
            If Info(1) > 0 And IsNumeric(Info(1)) Then
                Get #FileNum, Info(1), chunk
                BytesReceived = Info(1)
            Else
                Get #FileNum, 1, chunk
            End If
            Socket.SendData chunk
            DoEvents
        Case "DFT"
            If Status <> "Requesting..." Then Exit Sub
            State "Connected..."
            Close #FileNum
        Case Else
            If Status = "Waiting..." Then
                BytesReceived = BytesReceived + bytesTotal
                State "Receiving..."
                Dim Dis As Long
                If LOF(FileNum) < 1 Then
                    Dis = 1
                Else
                    Dis = LOF(FileNum)
                End If
                Put #FileNum, Dis, Sata
                Checkit
            ElseIf Status = "Receiving..." Then
                BytesReceived = BytesReceived + bytesTotal
                Put #FileNum, , Sata
                Checkit
            End If
    End Select
End Sub

Private Sub Checkit()
    If LOF(FileNum) > FileSize - 1 Then
        Close #FileNum
        tmrInfo_Timer
        State "Connected..."
        MsgBox "File Transfer Complete [" & Filename & "]", , "Transfer Complete"
        cmdSend.Enabled = True
    End If
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Socket_Close
End Sub

Private Sub State(State As String)
    DoEvents
    If State = "Disconnected..." Then
        cmdSend.Enabled = False
        cmdListen.Enabled = True
        cmdConnect.Enabled = True
    ElseIf State = "Connecting..." Or State = "Listening..." Or State = "Connected..." Then
        If State = "Connected..." Then cmdSend.Enabled = True
        cmdListen.Enabled = False
        cmdConnect.Enabled = False
    End If
    Status = State
    lblStatus.Caption = Status
End Sub

Private Sub Socket_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    On Error Resume Next
    If Status = "Sending..." Then
    DoEvents
    SSF = SSF + bytesSent
        If SSF >= Chunka - 1 Then
            Dim chunk As String
            Chunka = 4096
            chunk = Space$(4096)
            BytesReceived = BytesReceived + bytesSent
            If FileSize - BytesReceived < 4096 Then
                chunk = Space$(FileSize - BytesReceived + 1)
            End If
            Get #FileNum, , chunk
            Socket.SendData chunk
            If BytesReceived > FileSize - 1 Then
                tmrInfo_Timer
                MsgBox "File Transfer Complete [" & Filename & "]", , "Transfer Complete"
                Close #FileNum
                State "Connected..."
                cmdSend.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub tmrInfo_Timer()
    On Error Resume Next
    Dim CHOLO As String
    If Status = "Receiving..." Or Status = "Sending..." Then
        CHOLO = BytesReceived
        If Status = "Receiving..." Then CHOLO = LOF(FileNum)
        Stat(0).Caption = "Filename : " & Filename
        Stat(1).Caption = "Size : " & DoConv(CHOLO) & " / " & DoConv(CStr(FileSize))
        Stat(2).Caption = "Speed : " & DoConv(CHOLO - LCHECK) & "/ps"
        Dim hour, min, sec
        sec = Int((FileSize - CHOLO) / ((CHOLO - LCHECK)))
        min = Int(sec / 60)
        hour = Int(min / 60)
        sec = Int(sec - min * 60)
        min = Int(min - hour * 60)
        LCHECK = BytesReceived
        P1.min = 0
        P1.Max = FileSize
        P1.Value = BytesReceived
        Stat(3).Caption = "Percent : " & dec(CHOLO / FileSize * 100, 2) & "%"
        Stat(4).Caption = "Time Remaining : " & hour & ":" & min & ":" & sec
    Else
        P1.Value = 0
        Stat(0).Caption = "Filename : "
        Stat(1).Caption = "Size : "
        Stat(2).Caption = "Speed : "
        Stat(3).Caption = "Percent : "
        Stat(4).Caption = "Time Remaining : "
    End If
End Sub

Private Function DoConv(Number As String)
Dim DoConv3 As String
DoConv3 = Number
Dim part As Variant
On Error Resume Next
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " B"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " KB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " MB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " GB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " TB"
        Exit Function
    End If
        If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " PB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " EB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " ZB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " YB"
        Exit Function
    End If
    DoConv = "? B"
End Function

Private Function dec(Number As String, lod As Integer)
Dim X As Variant
X = Split(Number, ".")
If UBound(X) < 1 Then
    dec = Number
Else
    If Len(X(1)) < lod Then
        dec = Number
    Else
        dec = X(0) & "." & Left(X(1), lod)
    End If
End If
End Function

