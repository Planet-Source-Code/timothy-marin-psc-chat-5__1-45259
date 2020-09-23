VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v5 Server"
   ClientHeight    =   4935
   ClientLeft      =   4170
   ClientTop       =   3270
   ClientWidth     =   5775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5775
   Begin VB.Timer tmrSecond 
      Interval        =   1000
      Left            =   5880
      Top             =   2400
   End
   Begin MSWinsockLib.Winsock WebSock 
      Index           =   0
      Left            =   5880
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   56
   End
   Begin MSWinsockLib.Winsock ChatSock 
      Index           =   0
      Left            =   5880
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   15003
   End
   Begin MSComctlLib.ImageList imgLToolbar 
      Left            =   4920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":014A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":04E4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":087E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C18
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FB2
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":134C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C9E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4038
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43D2
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":476C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1005
      ButtonWidth     =   1164
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgLToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Banned"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Restart"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgStatus 
      Left            =   5880
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":523A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":596E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":629A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6634
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7102
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":749C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D736
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmNav 
      Caption         =   "Users"
      Height          =   4335
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      Begin MSComctlLib.ListView LVU 
         Height          =   3975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgStatus"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Address"
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Level"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Muzzle"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Timout"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Clear"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Repeat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Last Msg"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Ammount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Idle"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Typing"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "PM Chat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Mail"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame frmNav 
      Caption         =   "Options"
      Height          =   4335
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   5775
      Begin VB.Frame frmTopic 
         Caption         =   "Topic"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   5535
         Begin VB.TextBox txtTopic 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Text            =   "Visual Basic"
            Top             =   240
            Width           =   5295
         End
      End
   End
   Begin VB.Frame frmNav 
      Caption         =   "Banned"
      Height          =   4335
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   5775
      Begin MSComctlLib.ListView LVB 
         Height          =   3975
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Address"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   2999
         EndProperty
      End
   End
   Begin VB.Frame frmNav 
      Caption         =   "Log"
      Height          =   4335
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   5775
      Begin SHDocVwCtl.WebBrowser Web 
         Height          =   3975
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5535
         ExtentX         =   9763
         ExtentY         =   7011
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Menu mnuB 
      Caption         =   "BAN"
      Visible         =   0   'False
      Begin VB.Menu mnuRemB 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu mnuRA 
         Caption         =   "Remove All"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "USER"
      Visible         =   0   'False
      Begin VB.Menu mnuKick 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status"
         Begin VB.Menu mnuSUSE 
            Caption         =   "User"
         End
         Begin VB.Menu mnuSOP 
            Caption         =   "Mod"
         End
         Begin VB.Menu mnuSAdmin 
            Caption         =   "Admin"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1
Private Sub ChatSock_Close(Index As Integer)
    On Error Resume Next
    AddBan "", ChatSock(Index).RemoteHostIP, 2
    AddLog "<b>" & ChatSock(Index).RemoteHostIP & " Has Disconnected...</b>"
    ChatSock(Index).Close
    If IsAuth(Index) <> 0 Then
        SendAllSocket "ChatSock", "part" & Chr(1) & LVU.ListItems("x" & Index), 0, Me
    End If
    LVU.ListItems.Remove LVU.ListItems("x" & Index).Index
    If Index = 0 Then ChatSock(0).Listen
End Sub

Function IsBan(IP As String) As Boolean
    On Error Resume Next
    IsBan = False
    Dim i As Integer
    For i = 1 To LVB.ListItems.Count
        If LVB.ListItems.Item(i).SubItems(1) = IP Then
            IsBan = True
            Exit Function
        End If
    Next
End Function
Sub AddBan(NM As String, IP As String, TT As Long)
    On Error Resume Next
    LVB.ListItems.Add , , NM
    LVB.ListItems.Item(LVB.ListItems.Count).SubItems(1) = IP
    LVB.ListItems.Item(LVB.ListItems.Count).SubItems(2) = TT
End Sub


Function NameInUse(Name As String) As String
    On Error Resume Next
    NameInUse = "0"
    For i = 1 To LVU.ListItems.Count
        If LCase(LVU.ListItems.Item(i)) = LCase(Name) Then
            NameInUse = Right(LVU.ListItems.Item(i).Key, Len(LVU.ListItems.Item(i).Key) - 1)
            Exit Function
        End If
    Next
End Function


Private Sub ChatSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'Check For Online
    'Check For Banned
    Dim i As Integer
    i = OpenSocket("ChatSock", Me)
    If IsBan(ChatSock(Index).RemoteHostIP) Then
        Exit Sub
    End If
    ChatSock(i).Accept requestID
    If IsOnline("ChatSock", ChatSock(i).RemoteHostIP, Me) > 3 Then
        ChatSock(i).Close
        Exit Sub
    End If
    AddLog "<b>" & ChatSock(i).RemoteHostIP & " Has Connected...</b>"
End Sub

Private Sub ChatSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Recieve Msg From Client
    On Error Resume Next
    Dim Sata As String
    If ChatSock(Index).State = 7 Then ChatSock(Index).GetData Sata
    DoEvents
    'Split Data
    LVU.ListItems("x" & Index).SubItems(5) = "0"
    DoEvents
    Set recSet = New Recordset
    Dim Stage1() As String '<-Split All Msgs
    Dim Stage2() As String '<-Split Header From Vars
    Dim Stage3() As String '<-Split Vars
    Stage1 = Split(Sata, Chr(0)) 'Complete Stage1
    'Loop Through Msgs and Parse
    Dim i As Integer
    Dim Xint As Long
    For i = 0 To UBound(Stage1)
        On Error Resume Next
        Stage2 = Split(Stage1(i), Chr(1)) 'Complete Stage2
        'Find Type Of Msg
        Select Case LCase(Stage2(0))
            Case "op"
                If UBound(Stage2) < 1 Then GoTo Errs 'Require Variables
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 1 Then GoTo Errs
                If LVU.ListItems("x" & Index).SubItems(3) = 0 Then Exit Sub 'USER NOT AN OP
                Select Case LCase(Stage3(0))
                    Case "kick"
                        If NameInUse(Stage3(1)) > 0 Then
                            If LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) >= LVU.ListItems("x" & Index).SubItems(3) Then Exit Sub
                            SendAllSocket "chatsock", "server" & Chr(1) & LVU.ListItems.Item("x" & Index) & " is kicking " & Stage3(1), 0, Me
                            ChatSock_Close NameInUse(Stage3(1))
                        End If
                    Case "ban"
                        If LVU.ListItems("x" & Index).SubItems(3) < 2 Then Exit Sub 'REQUIRE ADMIN STATUS
                        If NameInUse(Stage3(1)) > 0 Then
                            If LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) >= LVU.ListItems("x" & Index).SubItems(3) Then Exit Sub
                            SendAllSocket "chatsock", "server" & Chr(1) & LVU.ListItems.Item("x" & Index) & " is banning " & Stage3(1), 0, Me
                            AddBan LCase(Stage3(1)), ChatSock(NameInUse(Stage3(1))).RemoteHostIP, 3600
                            ChatSock_Close NameInUse(Stage3(1))
                        End If
                    Case "mod"
                        If LVU.ListItems("x" & Index).SubItems(3) < 2 Then Exit Sub 'REQUIRE ADMIN STATUS
                        If NameInUse(Stage3(1)) > 0 Then
                            If LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) >= LVU.ListItems("x" & Index).SubItems(3) Then Exit Sub
                            SendAllSocket "chatsock", "server" & Chr(1) & LVU.ListItems.Item("x" & Index) & " is setting " & Stage3(1) & " as a moderator", 0, Me
                            LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) = 1
                            recSet.Close
                            recSet.Open "Select * from Users Where Name=""" & LCase(Stage3(1)) & """", dbConn, 1, 3
                            recSet("level") = 1
                            recSet.Update
                            recSet.Close
                        End If
                    Case "dmod"
                        If LVU.ListItems("x" & Index).SubItems(3) < 2 Then Exit Sub 'REQUIRE ADMIN STATUS
                        If NameInUse(Stage3(1)) > 0 Then
                            If LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) >= LVU.ListItems("x" & Index).SubItems(3) Then Exit Sub
                            SendAllSocket "chatsock", "server" & Chr(1) & LVU.ListItems.Item("x" & Index) & " is setting " & Stage3(1) & " as a user", 0, Me
                            LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) = 0
                            recSet.Close
                            recSet.Open "Select * from Users Where Name=""" & LCase(Stage3(1)) & """", dbConn, 1, 3
                            recSet("level") = 0
                            recSet.Update
                            recSet.Close
                        End If
                    Case "admin"
                        If LVU.ListItems("x" & Index).SubItems(3) < 2 Then Exit Sub 'REQUIRE ADMIN STATUS
                        If NameInUse(Stage3(1)) > 0 Then
                            If LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) >= LVU.ListItems("x" & Index).SubItems(3) Then Exit Sub
                            SendAllSocket "chatsock", "server" & Chr(1) & LVU.ListItems.Item("x" & Index) & " is setting " & Stage3(1) & " as an Admin", 0, Me
                            LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) = 2
                            recSet.Close
                            recSet.Open "Select * from Users Where Name=""" & LCase(Stage3(1)) & """", dbConn, 1, 3
                            recSet("level") = 2
                            recSet.Update
                            recSet.Close
                        End If
                    Case "cab"
                        If LVU.ListItems("x" & Index).SubItems(3) < 2 Then Exit Sub 'REQUIRE ADMIN STATUS
                        SendAllSocket "chatsock", "server" & Chr(1) & LVU.ListItems.Item("x" & Index) & " is clearing all current bans.", 0, Me
                        LVB.ListItems.Clear
                    Case "muzzle"
                        If NameInUse(Stage3(1)) > 0 Then
                            If LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(3) >= LVU.ListItems("x" & Index).SubItems(3) Then Exit Sub
                            SendAllSocket "chatsock", "server" & Chr(1) & LVU.ListItems.Item("x" & Index) & " is muzzling " & Stage3(1) & " for 1 minute", 0, Me
                            LVU.ListItems("x" & NameInUse(Stage3(1))).SubItems(4) = "60"
                        End If
                End Select
            Case Chr(5)
                If LVU.ListItems("x" & Index).SubItems(11) < 2 Then
                    LVU.ListItems("x" & Index).SubItems(11) = "3"
                    LVU.ListItems("x" & Index).ForeColor = vbBlue
                    SendAllSocket "ChatSock", "tm" & Chr(1) & LVU.ListItems("x" & Index), 0, Me
                End If
            Case Chr(7) 'Name PM Typing Status
                If UBound(Stage2) < 1 Then GoTo Errs
                If LVU.ListItems("x" & Index).SubItems(11) < 2 Then
                    LVU.ListItems("x" & Index).SubItems(11) = "3"
                    LVU.ListItems("x" & Index).ForeColor = vbRed
                    LVU.ListItems("x" & Index).SubItems(12) = Stage2(1)
                    'MsgBox NameInUse(LCase(Stage2(1)))
                    SendSocket "ChatSock", "pm" & Chr(1) & LVU.ListItems("x" & Index), NameInUse(LCase(Stage2(1))), Me
                End If
            Case "create" 'Name,Password,Email,WebSite,Skills,Picture,Country,Age
                On Error GoTo CreateErr
                If frmMain.IsAuth(Index) <> 0 Then
                    SendSocket "ChatSock", "auth" & Chr(1) & "1", Index, Me
                End If
                If UBound(Stage2) < 1 Then
                    'Tell User Packet Error
                    SendSocket "ChatSock", "auth" & Chr(1) & "1", Index, Me
                    DoEvents
                    GoTo Errs 'Require Variables
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 7 Then 'Require 8 Variables
                    'Tell User Packet Error
                    SendSocket "ChatSock", "auth" & Chr(1) & "1", Index, Me
                    DoEvents
                    GoTo Errs
                End If
                If Len(Stage3(0)) > 15 Or Len(Stage3(0)) < 3 Or Len(Stage3(1)) > 15 Or Len(Stage3(1)) < 3 Then
                    'Tell User Name or pass is invalid size
                    SendSocket "ChatSock", "auth" & Chr(1) & "2", Index, Me
                    DoEvents
                    GoTo Errs
                End If
                For Xint = 0 To UBound(Stage3)
                    If IChr(Stage3(Xint)) <> Stage3(Xint) Or InStr(Stage3(0), Chr(32)) > 1 Or InStr(Stage3(1), Chr(32)) > 1 Then
                    'Tell User he used invalid Charicters
                    SendSocket "ChatSock", "auth" & Chr(1) & "3", Index, Me
                    DoEvents
                    GoTo Errs
                    End If
                Next
                recSet.Open "Select * from Users where address=""" & ChatSock(Index).RemoteHostIP & """", dbConn, 1, 3
                If recSet.RecordCount > 3 Then
                    recSet.Close
                    'Tell user Name in use// This is bull Really the user has to many accounts on that ip
                    SendSocket "ChatSock", "auth" & Chr(1) & "4", Index, Me
                    DoEvents
                    GoTo Errs
                End If
                recSet.Close
                recSet.Open "Select * from Users where Name=""" & LCase(Stage3(0)) & """", dbConn, 1, 3
                If recSet.RecordCount > 0 Then
                    recSet.Close
                    'Tell user Name in use
                    SendSocket "ChatSock", "auth" & Chr(1) & "4", Index, Me
                    DoEvents
                    GoTo Errs
                End If
                'Everything Looks Ok Add The Account
                recSet.AddNew
                recSet("name") = LCase(Stage3(0))
                recSet("password") = Stage3(1)
                recSet("email") = Stage3(2)
                recSet("website") = Stage3(3)
                recSet("skills") = Stage3(4)
                recSet("picture") = Stage3(5)
                recSet("country") = Stage3(6)
                recSet("address") = ChatSock(Index).RemoteHostIP
                recSet("age") = Stage3(7)
                recSet("created") = Now
                recSet("lastlogin") = Now
                recSet("lastlogout") = Now
                recSet("level") = "0"
                recSet("muzzle") = "0"
                recSet.Update
                recSet.Close
                Set recSet = Nothing
                'Tell Channel Account Created
                SendAllSocket "ChatSock", "server" & Chr(1) & Now & " : " & Stage3(0) & " Has Registered Account. ", 0, Me
                'Log User in
                LoginUser Stage3(0), Index
                Exit Sub
                GoTo Errs
CreateErr:
                MsgBox Err.Description
                'Tell User Unknown Login Error
                SendSocket "ChatSock", "auth" & Chr(1) & "11", Index, Me
                DoEvents
                GoTo Errs
            Case "login" 'Name,Password
                If frmMain.IsAuth(Index) <> 0 Then
                    SendSocket "ChatSock", "auth" & Chr(1) & "1", Index, Me
                End If
                On Error GoTo LoginErr
                If UBound(Stage2) < 1 Then
                    SendSocket "ChatSock", "auth" & Chr(1) & "5", Index, Me
                    DoEvents
                    GoTo Errs 'Require Variables
                    'Tell User Packet Error
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 1 Then 'Require 2 Variables
                    SendSocket "ChatSock", "auth" & Chr(1) & "5", Index, Me
                    DoEvents
                    GoTo Errs
                    'Tell User Packet Error
                End If
                If Len(Stage3(0)) > 15 Or Len(Stage3(0)) < 3 Or Len(Stage3(1)) > 15 Or Len(Stage3(1)) < 3 Then
                    'Tell User Name or pass is invalid size
                    SendSocket "ChatSock", "auth" & Chr(1) & "6", Index, Me
                    DoEvents
                    GoTo Errs
                End If
                For Xint = 0 To UBound(Stage3)
                    If IChr(Stage3(Xint)) <> Stage3(Xint) Then
                    'Tell User he used invalid Charicters
                    SendSocket "ChatSock", "auth" & Chr(1) & "7", Index, Me
                    DoEvents
                    GoTo Errs
                    End If
                Next
                'At this Stage We know its a valid name and pass. so check against db
                Set recSet = New Recordset
                recSet.Open "Select * from Users where Name=""" & LCase(Stage3(0)) & """ and password = """ & Stage3(1) & """", dbConn, 1, 3
                If recSet.RecordCount < 1 Then
                    recSet.Close
                    'Tell User Account Not Found
                    SendSocket "ChatSock", "auth" & Chr(1) & "8", Index, Me
                    DoEvents
                    GoTo Errs
                End If
                'Check if account is in use (9)
                If Int(NameInUse(Stage3(0))) > 0 Then
                    SendSocket "ChatSock", "auth" & Chr(1) & "9", Index, Me
                    DoEvents
                    GoTo Errs
                End If
                'OtherWize Allow Login
                LoginUser Stage3(0), Index
                Exit Sub
                GoTo Errs
LoginErr:
                'Tell User Unknown Login Error
                SendSocket "ChatSock", "auth" & Chr(1) & "10", Index, Me
                DoEvents
                GoTo Errs
            Case "flash" 'Flash client uses it to set its status
                    SendAllSocket "ChatSock", "status" & Chr(1) & LVU.ListItems("x" & Index) & Chr(2) & "12", 0, Me
                    LVU.ListItems("x" & Index).SmallIcon = 12
            Case "status" '1-7
                '1 online | 3 away | 4 brb | 6 coding | 7 eating | 8 phone <- Codes Client Sends
                If UBound(Stage2) < 1 Then GoTo Errs
                    'Show Error
                    Select Case Stage2(1)
                        Case "1", "3", "4", "5", "6", "7", "8"
                            If StatusToInt("1", Stage2(1)) > 0 Then
                                SendAllSocket "ChatSock", "status" & Chr(1) & LVU.ListItems("x" & Index) & Chr(2) & StatusToInt("1", Stage2(1)), 0, Me
                                LVU.ListItems("x" & Index).SmallIcon = Int(StatusToInt("1", Stage2(1)))
                            End If
                    End Select
            Case "msg", "pmsg" 'Font|Size|Bold|Italic|Underline|Color|Msg (|USER)
            LVU.ListItems("x" & Index).SubItems(10) = "0"
            If LVU.ListItems("x" & Index).SmallIcon = 7 Then
                SendAllSocket "ChatSock", "status" & Chr(1) & LVU.ListItems("x" & Index) & Chr(2) & "1", 0, Me
                LVU.ListItems("x" & Index).SmallIcon = 1
            End If
            'MsgBox Sata
                If LVU.ListItems("x" & Index).SubItems(4) > 0 Then Exit Sub
                If UBound(Stage2) < 1 Then
                    GoTo Errs 'Require Variables
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 6 Then 'Require 7 Variables
                    GoTo Errs
                End If
                Stage3(6) = Replace(Stage3(6), Chr(13), "<br>")
                Stage3(6) = Replace(Stage3(6), Chr(9), ":tab")
                Stage3(6) = IChr(Stage3(6))
                If Len(Stage3(6)) > 351 Or Len(Stage3(6)) < 1 Then GoTo Errs
                'ONLY ALLOW 6 LINE BREAKS
                Dim ML() As String
                ML = Split(Stage3(6), "&lt;br&gt;")
                Stage3(6) = ""
                For Xint = 0 To UBound(ML) - 1
                    If Xint < 6 Then
                        Stage3(6) = Stage3(6) & ML(Xint) & "<br>"
                    Else
                        Stage3(6) = Stage3(6) & " " & ML(Xint)
                    End If
                Next
                Stage3(6) = Stage3(6) & " " & ML(UBound(ML))
                'END LINE BREAK CHECK
                If LVU.ListItems("x" & Index).SubItems(8) = Stage3(6) Then
                    LVU.ListItems("x" & Index).SubItems(7) = LVU.ListItems("x" & Index).SubItems(7) + 1
                    If LVU.ListItems("x" & Index).SubItems(7) > 2 Then
                        SendAllSocket "ChatSock", "server" & Chr(1) & "User " & LVU.ListItems("x" & Index) & " Was Kicked By Server(Repeat).", 0, Me
                        ChatSock_Close Index
                        Exit Sub
                    End If
                End If
                LVU.ListItems("x" & Index).SubItems(8) = Stage3(6)
                LVU.ListItems("x" & Index).SubItems(9) = LVU.ListItems("x" & Index).SubItems(9) + 1
                If LVU.ListItems("x" & Index).SubItems(9) > 9 Then
                        SendAllSocket "ChatSock", "server" & Chr(1) & "User " & LVU.ListItems("x" & Index) & " Was Kicked By Server(Flood).", 0, Me
                        ChatSock_Close Index
                        Exit Sub
                    End If
                Dim SMsg As String
                SMsg = Stage3(0)
                For Xint = 1 To 6
                    SMsg = SMsg & Chr(2) & Stage3(Xint)
                Next
                
                If LCase(Stage2(0)) = "msg" Then
                    SendAllSocket "ChatSock", "msg" & Chr(1) & SMsg & Chr(2) & LVU.ListItems("x" & Index), 0, Me
                Else
                    If UBound(Stage3) < 7 Then 'Require 8 Variables
                        GoTo Errs
                    End If
                    SendSocket "ChatSock", "pmsg" & Chr(1) & SMsg & Chr(2) & LVU.ListItems("x" & Index) & Chr(2) & Stage3(7), Index, Me
                    SendSocket "ChatSock", "pmsg" & Chr(1) & SMsg & Chr(2) & LVU.ListItems("x" & Index), NameInUse(LCase(Stage3(7))), Me
                End If
            Case "smail" 'User,Subject,Message
                If UBound(Stage2) < 1 Then
                    GoTo Errs 'Require Variables
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 2 Then 'Require 3 Variables
                    GoTo Errs
                End If
                If LVU.ListItems("x" & Index).SubItems(13) > 0 Then
                    SendSocket "ChatSock", "auth" & Chr(1) & "16", Index, Me 'Send Error Invalid Mail no User
                    DoEvents
                    GoTo Errs:
                End If
                recSet.Open "Select * from Users where Name=""" & LCase(Stage3(0)) & """", dbConn, 1, 3
                    If recSet.RecordCount < 1 Then
                        SendSocket "ChatSock", "auth" & Chr(1) & "15", Index, Me 'Send Error Invalid Mail no User
                        DoEvents
                        GoTo Errs:
                    End If
                    LVU.ListItems("x" & Index).SubItems(13) = "45"
                recSet.Close
                recSet.Open "Select * from Mail where To=""" & LCase(Stage3(0)) & """", dbConn, 1, 3
                        recSet.AddNew
                        recSet("from") = LCase(LVU.ListItems("x" & Index))
                        recSet("to") = LCase(Stage3(0))
                        recSet("subject") = Stage3(1)
                        recSet("message") = Stage3(2)
                        recSet("date") = Now
                        recSet.Update
                        SendSocket "ChatSock", "inbox" & Chr(1) & recSet.RecordCount, NameInUse(LCase(Stage3(0))), Me
                recSet.Close
                DoEvents
                SendSocket "ChatSock", "newmail" & Chr(1) & LVU.ListItems("x" & Index), Int(NameInUse(LCase(Stage3(0)))), Me
                DoEvents
                SendSocket "ChatSock", "sent" & Chr(1), Index, Me 'Tell Sender it was succesfull
            Case "dmail" 'Delete Mail
                If UBound(Stage2) < 1 Then GoTo Errs
                dbConn.Execute "DELETE FROM MAIL WHERE TO='" & LCase(LVU.ListItems("x" & Index)) & "' AND KEY=" & Int(Stage2(1))
                recSet.Close
                recSet.Open "Select * from mail where To=""" & LCase(LVU.ListItems("x" & Index)) & """", dbConn, 1, 3
                SendSocket "ChatSock", "inbox" & Chr(1) & recSet.RecordCount, Index, Me
                recSet.Close
                
            Case "cmail" 'Check Mail
                recSet.Close
                recSet.Open "Select * from mail where To=""" & LCase(LVU.ListItems("x" & Index)) & """", dbConn, 1, 3
                If recSet.RecordCount > 0 Then
                    recSet.MoveFirst
                    Do Until recSet.EOF
                        DoEvents
                        SendSocket "ChatSock", "mail" & Chr(1) & recSet("key") & Chr(2) & recSet("date") & Chr(2) & recSet("from") & Chr(2) & recSet("subject") & Chr(2) & recSet("message"), Index, Me
                        recSet.MoveNext
                    Loop
                End If
                SendSocket "ChatSock", "endmail" & Chr(1) & recSet.RecordCount, Index, Me
                recSet.Close
            Case "profile"
                If UBound(Stage2) < 1 Then GoTo Errs
                'MsgBox Int(NameInUse(Stage2(1)))
                SendSocket "ChatSock", "profile" & Chr(1) & Stage2(1) & Chr(2) & LVU.ListItems.Item("x" & NameInUse(Stage2(1))).Tag, Index, Me
                
            Case "ep"
                If UBound(Stage2) < 1 Then
                    'Tell User Packet Error
                    GoTo Errs 'Require Variables
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 5 Then 'Require 6 Variables
                    'Tell User Packet Error
                    GoTo Errs
                End If
                recSet.Close
                recSet.Open "Select * from Users where Name=""" & LCase(LVU.ListItems("x" & Index)) & """", dbConn, 1, 3
                'Everything Looks Ok Add The Account
                If recSet.RecordCount > 0 Then
                recSet.MoveFirst
                recSet("email") = Stage3(0)
                recSet("website") = Stage3(1)
                recSet("skills") = Stage3(2)
                recSet("picture") = Stage3(3)
                recSet("country") = Stage3(4)
                recSet("address") = ChatSock(Index).RemoteHostIP
                recSet("age") = Stage3(5)
                
                LVU.ListItems("x" & Index).Tag = recSet("email") & Chr(2) & _
                                            recSet("website") & Chr(2) & _
                                            recSet("skills") & Chr(2) & _
                                            recSet("picture") & Chr(2) & _
                                            recSet("country") & Chr(2) & _
                                            recSet("age") & Chr(2) & _
                                            recSet("created") & Chr(2) & _
                                            recSet("lastlogin") & Chr(2) & _
                                            recSet("lastlogout") & Chr(2) & _
                                            recSet("level")
                recSet.Update
                recSet.Close
                End If
        End Select
Errs:
    Next
End Sub

Function IsAuth(Index As Integer)
    On Error Resume Next
    IsAuth = 0
    IsAuth = LVU.ListItems("x" & Index).Index
    'MsgBox IsAuth
End Function

Sub LoginUser(Name As String, Index As Integer)
    'Send Him Userlist
    On eror GoTo shoot:
    Dim i As Integer
    SendSocket "ChatSock", "topic" & Chr(1) & txtTopic.Text, Index, Me
    DoEvents
    For i = 1 To LVU.ListItems.Count
        SendSocket "ChatSock", "list" & Chr(1) & LVU.ListItems.Item(i) & Chr(2) & LVU.ListItems.Item(i).SmallIcon, Index, Me
    Next
        Set recSet = New Recordset
        recSet.Open "Select * from Users where Name=""" & LCase(Name) & """", dbConn, 1, 3
        LVU.ListItems.Add , "x" & Index, Name, , 1
        LVU.ListItems("x" & Index).SubItems(1) = ChatSock(Index).RemoteHostIP
        LVU.ListItems("x" & Index).SubItems(2) = Now
        LVU.ListItems("x" & Index).SubItems(3) = recSet("level")
        LVU.ListItems("x" & Index).SubItems(4) = recSet("muzzle")
        LVU.ListItems("x" & Index).SubItems(5) = "0"
        LVU.ListItems("x" & Index).SubItems(6) = "0"
        LVU.ListItems("x" & Index).SubItems(7) = "0"
        LVU.ListItems("x" & Index).SubItems(8) = ""
        LVU.ListItems("x" & Index).SubItems(9) = "0"
        LVU.ListItems("x" & Index).SubItems(10) = "0"
        LVU.ListItems("x" & Index).Tag = recSet("email") & Chr(2) & _
                                            recSet("website") & Chr(2) & _
                                            recSet("skills") & Chr(2) & _
                                            recSet("picture") & Chr(2) & _
                                            recSet("country") & Chr(2) & _
                                            recSet("age") & Chr(2) & _
                                            recSet("created") & Chr(2) & _
                                            recSet("lastlogin") & Chr(2) & _
                                            recSet("lastlogout") & Chr(2) & _
                                            recSet("level")
    recSet.Close
    recSet.Open "Select * from Mail where To=""" & LCase(Name) & """", dbConn, 1, 3
    SendSocket "ChatSock", "Inbox" & Chr(1) & recSet.RecordCount, Index, Me
    recSet.Close
    Set recSet = Nothing
      'Send All Login
      SendAllSocket "ChatSock", "join" & Chr(1) & Name & Chr(2) & "1", 0, Me
Exit Sub
shoot:
    ChatSock_Close Index
End Sub


Function StatusToInt(hmm As String, Hmmm As String)
    'online phone away brb coding eating idle 12-flash 13-muzzle
    '1 online 2 idle 3 away 4 brb 5 flash 6 coding 7 eating 8 phone <- codes u send client
    If hmm = "1" Then 'You want to go Client to Imglist
        StatusToInt = 0
        Select Case Hmmm
            Case "1"
                StatusToInt = 1
            Case "2"
                StatusToInt = 7
            Case "3"
                StatusToInt = 3
            Case "4"
                StatusToInt = 4
            Case "5"
                StatusToInt = 12
            Case "6"
                StatusToInt = 5
            Case "7"
                StatusToInt = 6
            Case "8"
                StatusToInt = 2
        End Select
    Else ' you want to go imagelist to client
        Select Case Hmmm
            Case "1"
                StatusToInt = 1
            Case "7"
                StatusToInt = 2
            Case "3"
                StatusToInt = 3
            Case "4"
                StatusToInt = 4
            Case "12"
                StatusToInt = 5
            Case "5"
                StatusToInt = 6
            Case "6"
                StatusToInt = 7
            Case "2"
                StatusToInt = 8
        End Select
    End If
End Function

Private Sub ChatSock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ChatSock_Close Index
End Sub

Private Sub Form_Load()
    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.ChangeIcon Me.Icon
    gSysTray.IconInSysTray
    On Error GoTo Error
    'ClearLog and navigate
    ClearLog
    'Connect To DataBase
    Set dbConn = New Connection
    dbConn.CursorLocation = adUseClient
    dbConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source = Accounts.mdb;"
    'Start Server
    ChatSock(0).Listen
    WebSock(0).Listen
    Exit Sub
Error:
    MsgBox "Error Loading Server : " & Err.Description, vbCritical
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim hmm As VbMsgBoxResult
    hmm = MsgBox("Are you sure you want to quit?", vbYesNo, "Quit")
    If hmm = vbYes Then
        gSysTray.RemoveFromSysTray
        Cancel = 0
    Else
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub LVB_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub LVB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Me.mnuB
End Sub

Private Sub LVU_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub LVU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Me.mnuUser
End Sub

Private Sub mnuBan_Click()
    On Error Resume Next
    AddBan LVU.SelectedItem, ChatSock(Int(Right(LVU.SelectedItem.Key, Len(LVU.SelectedItem.Key) - 1))).RemoteHostIP, 3600
    ChatSock_Close Int(Right(LVU.SelectedItem.Key, Len(LVU.SelectedItem.Key) - 1))
End Sub

Private Sub mnuKick_Click()
    'On Error Resume Next
    ChatSock_Close CInt(Right(LVU.SelectedItem.Key, CLng(Len(LVU.SelectedItem.Key) - 1)))
End Sub

Private Sub mnuRA_Click()
    LVB.ListItems.Clear
End Sub

Private Sub mnuRemB_Click()
    On Error Resume Next
    LVB.ListItems.Remove LVB.SelectedItem.Index
End Sub

Private Sub mnuSAdmin_Click()
On Error Resume Next
    Set recSet = New Recordset
    LVU.SelectedItem.SubItems(3) = 2
    recSet.Close
    recSet.Open "Select * from Users Where Name=""" & LCase(LVU.SelectedItem) & """", dbConn, 1, 3
    recSet("level") = 2
    recSet.Update
    recSet.Close
End Sub

Private Sub mnuSOP_Click()
On Error Resume Next
    Set recSet = New Recordset
    LVU.SelectedItem.SubItems(3) = 1
    recSet.Close
    recSet.Open "Select * from Users Where Name=""" & LCase(LVU.SelectedItem) & """", dbConn, 1, 3
    recSet("level") = 1
    recSet.Update
    recSet.Close
End Sub

Private Sub mnuSUSE_Click()
    On Error Resume Next
    Set recSet = New Recordset
    LVU.SelectedItem.SubItems(3) = 0
    recSet.Close
    recSet.Open "Select * from Users Where Name=""" & LCase(LVU.SelectedItem) & """", dbConn, 1, 3
    recSet("level") = 0
    recSet.Update
    recSet.Close
End Sub

Private Sub tmrSecond_Timer()
    On Error Resume Next
    Dim i As Integer
    For i = 1 To LVB.ListItems.Count
        LVB.ListItems.Item(i).SubItems(2) = LVB.ListItems.Item(i).SubItems(2) - 1
        If LVB.ListItems.Item(i).SubItems(2) < 1 Then
            LVB.ListItems.Remove i
        End If
    Next
    For i = 1 To LVU.ListItems.Count
        If LVU.ListItems.Item(i).SubItems(4) > 0 Then
            LVU.ListItems.Item(i).SubItems(4) = LVU.ListItems.Item(i).SubItems(4) - 1
        End If
        If LVU.ListItems.Item(i).SubItems(13) = "" Then LVU.ListItems.Item(i).SubItems(13) = 0
        If LVU.ListItems.Item(i).SubItems(13) > 0 Then
            LVU.ListItems.Item(i).SubItems(13) = LVU.ListItems.Item(i).SubItems(13) - 1
        End If
        LVU.ListItems.Item(i).SubItems(5) = LVU.ListItems.Item(i).SubItems(5) + 1
        If LVU.ListItems.Item(i).SubItems(5) > 120 Then
            ChatSock_Close Int(Right(LVU.ListItems.Item(i).Key, Len(LVU.ListItems.Item(i).Key) - 1))
            LVU.ListItems.Item(i).SubItems(5) = 0
        End If
        LVU.ListItems.Item(i).SubItems(6) = LVU.ListItems.Item(i).SubItems(6) + 1
        If LVU.ListItems.Item(i).SubItems(6) > 30 Then
            LVU.ListItems.Item(i).SubItems(6) = "0"
            LVU.ListItems.Item(i).SubItems(7) = "0"
            LVU.ListItems.Item(i).SubItems(8) = ""
            LVU.ListItems.Item(i).SubItems(9) = "0"
        End If
        LVU.ListItems.Item(i).SubItems(10) = LVU.ListItems.Item(i).SubItems(10) + 1
        If Int(LVU.ListItems.Item(i).SubItems(10)) > Int(300) Then
            LVU.ListItems.Item(i).SubItems(10) = 0
            If LVU.ListItems.Item(i).SmallIcon = 1 Then
                SendAllSocket "ChatSock", "status" & Chr(1) & LVU.ListItems.Item(i) & Chr(2) & "7", 0, Me
                LVU.ListItems.Item(i).SmallIcon = 7
            End If
        End If
        If LVU.ListItems.Item(i).SubItems(11) = "" Then LVU.ListItems.Item(i).SubItems(11) = "0"
        If LVU.ListItems.Item(i).SubItems(11) > -1 Then LVU.ListItems.Item(i).SubItems(11) = LVU.ListItems.Item(i).SubItems(11) - 1
            If Int(LVU.ListItems.Item(i).SubItems(11)) = 0 Then
            If LVU.ListItems.Item(i).SubItems(12) <> "" Then
                SendSocket "ChatSock", "ps" & Chr(1) & LVU.ListItems.Item(i), NameInUse(LCase(LVU.ListItems.Item(i).SubItems(12))), Me
                LVU.ListItems.Item(i).SubItems(12) = ""
                LVU.ListItems.Item(i).ForeColor = vbBlack
                DoEvents
                GoTo nexteh
            End If
            SendAllSocket "ChatSock", "st" & Chr(1) & LVU.ListItems.Item(i), 0, Me
            LVU.ListItems.Item(i).ForeColor = vbBlack
        End If
nexteh:
    Next
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 5 Then
        MsgBox "Writen By: Timothy Marin and Carsten Dressler" & vbCrLf & "www.intradream.com", , "v5"
    End If
    If Button.Index = 7 Then
        Dim i As Integer
        For i = 0 To ChatSock.UBound
            ChatSock_Close i
        Next
        For i = 0 To WebSock.UBound
            WebSock_Close i
        Next
    End If
    If Button.Index < 5 Then
        For i = 0 To 4
            frmNav(i).Visible = False
        Next
        frmNav(Button.Index).Visible = True
    End If
End Sub

Private Sub Web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Web.Document.body.scrolltop = Len(Web.Document.body.innertext)
End Sub

Function AddLog(msg As String)
    On Error Resume Next
    Web.Document.body.innerHTML = Web.Document.body.innerHTML & msg & "<br>"
    Web.Document.body.scrolltop = CLng(Len(Web.Document.body.innerHTML)) * 100
End Function

Function ClearLog()
'On Error Resume Next
    Web.Navigate "about:blank"
    Do While Web.ReadyState <> READYSTATE_COMPLETE
      DoEvents
    Loop
    DoEvents
    Web.Document.body.topmargin = 5
    Web.Document.body.leftmargin = 3
    Web.Document.body.innerHTML = "<b><i><font color=""#999999"">PSC Chat V5 Server Build 22</font></b></i><br>"
End Function

Private Sub WebSock_Close(Index As Integer)
    WebSock(Index).Close
    If Index = 0 Then WebSock(Index).Listen
End Sub

Private Sub WebSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'Check For Online
    'Check For Banned
    Dim i As Integer
    i = OpenSocket("WebSock", Me)
    If IsBan(WebSock(Index).RemoteHostIP) Then
        Exit Sub
    End If
    WebSock(i).Accept requestID
    If IsOnline("WebSock", WebSock(i).RemoteHostIP, Me) > 3 Then
        ChatSock(i).Close
        Exit Sub
    End If
    AddLog "<b>" & WebSock(i).RemoteHostIP & " Has Connected...</b>"
End Sub

Private Sub WebSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Sata As String
    Dim FFile As Integer
    Dim Fchunk As String
    Dim SWFSize As Long
    FFile = FreeFile
    WebSock(Index).GetData Sata
    DoEvents
    'Check the Get Statement for the SWF
    If Left(Sata, 3) = "GET" And InStr(LCase(Sata), LCase("Client.swf")) Then
        REPLY = "HTTP/1.1 200 OK" & vbCrLf & _
                "Server: Microsoft -IIS / 5.1" & vbCrLf & _
                "Date: " & Format(Date, "ddd, dd mmm yyyy") & " " & Format(Time, "HH:MM:SS") & " GMT" & vbCrLf & _
                "Content-Type: application/x-shockwave-flash" & vbCrLf & _
                "Accept -Ranges: bytes" & vbCrLf & _
                "Last-Modified:" & Format(Date, "ddd, dd mmm yyyy") & " " & Format(Time, "HH:MM:SS") & " GMT" & vbCrLf & _
                "Content-Length: SIZEOFFILE" & vbCrLf & vbCrLf
        FFile = FreeFile
        Open App.Path & "\Client.swf" For Binary As #FFile
            SWFSize = LOF(FFile)
            Fchunk = Space$(SWFSize)
            Get #FFile, , Fchunk
        Close FFile
        WebSock(Index).SendData Replace(REPLY, "SIZEOFFILE", SWFSize) & Fchunk
    Else
        WebSock_Close Index
    End If
End Sub

Private Sub WebSock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WebSock_Close Index
End Sub

Private Sub WebSock_SendComplete(Index As Integer)
WebSock_Close Index
End Sub
