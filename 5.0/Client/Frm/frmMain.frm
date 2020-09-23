VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "IntraDream PSC Chat v5"
   ClientHeight    =   6435
   ClientLeft      =   3630
   ClientTop       =   2700
   ClientWidth     =   7785
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   7785
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7200
      Top             =   2040
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   5640
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ImageList imgStatus 
      Left            =   7320
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0582
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":091C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1050
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":197C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D16
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":220A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":293E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A3F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A846
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11D48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrTimeout 
      Interval        =   60000
      Left            =   7320
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   5280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LVU 
      Height          =   4695
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "imgStatus"
      SmallIcons      =   "imgStatus"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Online"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBF 
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   4680
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons(3)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Font"
            ImageKey        =   "Small Caps"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Color"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Background"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   3
      Left            =   6360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":185AA
            Key             =   "Small Caps"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":186BC
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":187CE
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":188E0
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":189F2
            Key             =   "Button"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19238
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":195D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A106
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar S 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   6120
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Offline..."
            TextSave        =   "Offline..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8096
            Text            =   "Topic :"
            TextSave        =   "Topic :"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Inbox : 0"
            TextSave        =   "Inbox : 0"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmSend 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   7575
      Begin VB.PictureBox p 
         Height          =   630
         Left            =   6720
         ScaleHeight     =   570
         ScaleWidth      =   735
         TabIndex        =   7
         Top             =   150
         Width           =   790
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            Height          =   570
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtChat 
         Height          =   615
         Left            =   75
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   150
         Width           =   6540
      End
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   4230
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   7461
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
      Location        =   ""
   End
   Begin MSComctlLib.Toolbar TBM 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgStatus"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Externally"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear Chat"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Sripting"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Status"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Online"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Away"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Brb"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Coding"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eating"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Phone"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   7320
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnudiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuMail 
         Caption         =   "Mail Box"
      End
      Begin VB.Menu mnuEditAccount 
         Caption         =   "Edit Profile"
      End
      Begin VB.Menu mnuDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "File Transfer"
      End
      Begin VB.Menu mnuWhite 
         Caption         =   "Whiteboard"
      End
      Begin VB.Menu mnudiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeSearch 
         Caption         =   "Code Search"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPops 
         Caption         =   "Popup Notify"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSound 
         Caption         =   "Sound Notify"
         Checked         =   -1  'True
      End
      Begin VB.Menu div6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScript 
         Caption         =   "Scripting"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDiv4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVIG 
         Caption         =   "Clear Ignored List"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPM 
      Caption         =   "PM"
      Visible         =   0   'False
      Begin VB.Menu mnuIgnore 
         Caption         =   "Ignore"
      End
      Begin VB.Menu mnuUnignore 
         Caption         =   "Unignore"
      End
      Begin VB.Menu mnudiv8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "View Profile"
      End
      Begin VB.Menu mnuOPM 
         Caption         =   "Private Message"
      End
      Begin VB.Menu mnudiv7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOPS 
         Caption         =   "Operators"
         Begin VB.Menu mnuKick 
            Caption         =   "Kick"
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Ban"
         End
         Begin VB.Menu mnuMuzzle 
            Caption         =   "Muzzle"
         End
         Begin VB.Menu mnudis8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUSTAT 
            Caption         =   "Status"
            Begin VB.Menu mnuMod 
               Caption         =   "Moderate"
            End
            Begin VB.Menu mnuDmod 
               Caption         =   "Demoderate"
            End
            Begin VB.Menu mnudiv9 
               Caption         =   "-"
            End
            Begin VB.Menu mnuADMIN 
               Caption         =   "Set as Admin(cannot undo)"
            End
         End
         Begin VB.Menu mnuCAB 
            Caption         =   "Clear All Bans"
         End
      End
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "mnuIcon"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnudiv10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExits 
         Caption         =   "Exit"
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
Private Sub cmdSend_Click()
    'Font|Size|Bold|Italic|Underline|Color|Msg
    txtChat.SetFocus
    If Len(txtChat) < 1 Or Sock.State <> 7 Then Exit Sub
    SockSend "msg" & Chr(1) & txtChat.Font & Chr(2) & txtChat.FontSize & Chr(2) & txtChat.FontBold & Chr(2) & txtChat.FontItalic & Chr(2) & txtChat.FontUnderline & Chr(2) & DectoWebCol(txtChat.ForeColor) & Chr(2) & txtChat.Text
    txtChat.Text = ""
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

Private Sub gSysTray_RButtonUP()
    PopupMenu Me.mnuIcon
End Sub

Function SaveHtml(WB As WebBrowser)
    Dim FP As String
    FP = InputBox("What Name!", "LOG", "HTML")
    If Len(FP) < 1 Then Exit Function
    Dim FF As Integer
    FF = FreeFile
    Open App.Path & "\" & FP & ".html" For Output As #FF
        Print #FF, "<body leftmargin=5 topmargin=5 style=""background:" & WB.Document.body.Style.background & """>" & WB.Document.body.innerHTML
    Close #FF
    MsgBox "Save Succesfull!"
End Function

Private Sub Form_Load()
ClearLog
Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.ChangeIcon Me.Icon
    gSysTray.IconInSysTray
    ClearLog
    DoEvents
    On Error Resume Next
    txtChat.ForeColor = vbBlack
    Dim i
    For i = 1 To 200
        Pavail(i) = True
    Next
    txtChat.ForeColor = vbBlack
End Sub

Public Sub ResetScript()
    frmMain.Script.Reset
    Script.AddObject "frmMain", frmMain, True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Me.Hide
    LVU.Left = Me.ScaleWidth - LVU.Width
    LVU.Height = Me.ScaleHeight - (frmSend.Height + S.Height + TBM.Height) + 100
    frmSend.Width = Me.ScaleWidth
    p.Left = Me.ScaleWidth - (cmdSend.Width + 120)
    txtChat.Width = Me.ScaleWidth - (cmdSend.Width + 250)
    TBF.Width = LVU.Left
    Web.Width = TBF.Width
    Web.Height = LVU.Height - TBF.Height
    frmSend.Top = Web.Height + TBF.Height + TBM.Height - 100
    TBF.Top = frmSend.Top - TBF.Height + 100
    Timer1.Enabled = False
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub LVU_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Function PMName(msg As String)
    On Error Resume Next
    Dim i
    PMName = 0
    For i = 1 To 200
        If Pavail(i) = False Then
            If LCase(PRoom(i).Caption) = LCase(msg) Then
                PMName = i
                Exit Function
            End If
        End If
    Next
End Function

Function NewRoom()
    On Error Resume Next
    Dim i
    NewRoom = 0
    For i = 1 To 200
        If Pavail(i) = True Then
            NewRoom = i
            Exit Function
        End If
    Next
End Function

Private Sub LVU_DblClick()
    On Error GoTo quits:
    Dim pmm As Integer
    If LVU.SelectedItem.Index > 0 Then
        pmm = PMName(LVU.SelectedItem)
        If pmm > 0 Then
            PRoom(pmm).Show
            PRoom(pmm).SetFocus
            PRoom(pmm).WindowState = vbNormal
        Else
            pmm = NewRoom
            PRoom(pmm).Show
            PRoom(pmm).Pnum = pmm
            PRoom(pmm).Caption = LVU.SelectedItem
            PRoom(pmm).LMsg = Now
            Pavail(pmm) = False
            PRoom(pmm).S.SimpleText = "Started " & PRoom(pmm).LMsg & "..."
        End If
    End If
quits:
End Sub

Private Sub LVU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPM
End Sub

Private Sub mnuAbout_Click()
    frmSplash.Show
    frmSplash.Timer1.Enabled = False
    frmSplash.Timer1.Interval = 10000
    frmSplash.Timer1.Enabled = True
End Sub

Private Sub mnuADMIN_Click()
    On Error Resume Next
    SockSend "op" & Chr(1) & "admin" & Chr(2) & LVU.SelectedItem
End Sub

Private Sub mnuBan_Click()
    On Error Resume Next
    SockSend "op" & Chr(1) & "ban" & Chr(2) & LVU.SelectedItem
End Sub

Private Sub mnuCAB_Click()
    On Error Resume Next
    SockSend "op" & Chr(1) & "cab" & Chr(2)
End Sub

Private Sub mnuCodeSearch_Click()
    Dim a As String
    a = InputBox("Search For...", "Search", "Timothy Marin")
    If Len(a) > 0 Then
        Shell "Explorer.exe ""http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=1&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria="" & a"
    End If
End Sub

Public Sub mnuConnect_Click()
    If mnuConnect.Caption = "Connect" Then
        mnuConnect.Caption = "Disconnect"
        Sock.Close
        Sock.Connect "Pscchat4.ods.org", 15003 'Pscchat4.ods.org
        TBM.Buttons.Item(1).Enabled = False
        TBM.Buttons.Item(2).Enabled = True
        S.Panels.Item(1).Text = "Connecting..."
    Else
        mnuConnect.Caption = "Connect"
        Sock.Close
        Unload frmLogin
        TBM.Buttons.Item(1).Enabled = True
        TBM.Buttons.Item(2).Enabled = False
        S.Panels.Item(1).Text = "Offline..."
        LVU.ListItems.Clear
        AddLog "<font Color=#990000>Disconnected from Server...</font>"
    End If
End Sub


Private Sub mnuDmod_Click()
    On Error Resume Next
    SockSend "op" & Chr(1) & "dmod" & Chr(2) & LVU.SelectedItem
End Sub

Private Sub mnuEditAccount_Click()
On Error Resume Next
    Me.SockSend "profile" & Chr(1) & UserName
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExits_Click()
    Unload Me
End Sub

Private Sub mnuIgnore_Click()
    On Error Resume Next
    List1.AddItem LVU.SelectedItem
    LVU.ListItems.Item(LVU.SelectedItem.Index).SmallIcon = 8
End Sub

Private Sub mnuKick_Click()
    On Error Resume Next
    SockSend "op" & Chr(1) & "kick" & Chr(2) & LVU.SelectedItem
End Sub

Private Sub mnuMail_Click()
    frmMail.Show
End Sub

Private Sub mnuMod_Click()
    On Error Resume Next
    SockSend "op" & Chr(1) & "mod" & Chr(2) & LVU.SelectedItem
End Sub

Private Sub mnuMuzzle_Click()
    On Error Resume Next
    SockSend "op" & Chr(1) & "muzzle" & Chr(2) & LVU.SelectedItem
End Sub

Private Sub mnuOPM_Click()
LVU_DblClick
End Sub

Private Sub mnuPops_Click()
    mnuPops.Checked = Not (mnuPops.Checked)
End Sub

Private Sub mnuProfile_Click()
    On Error Resume Next
    Me.SockSend "profile" & Chr(1) & LVU.SelectedItem.Text
End Sub

Private Sub mnuScript_Click()
    mnuScript.Checked = Not (mnuScript.Checked)
End Sub

Private Sub mnuShow_Click()
Me.WindowState = vbNormal
Me.Show
End Sub

Private Sub mnuSound_Click()
    mnuSound.Checked = Not (mnuSound.Checked)
End Sub

Private Sub mnuTransfer_Click()
    Dim X As New frmTransfer
    X.Show
End Sub

Private Sub mnuUnignore_Click()
    On Error Resume Next
    If Len(LVU.SelectedItem.Text) < 0 Then Exit Sub
    LVU.ListItems.Item(LVU.SelectedItem.Index).SmallIcon = Int(LVU.ListItems.Item(LVU.SelectedItem.Index).Tag)
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
        If LCase(List1.List(i)) = LCase(LVU.SelectedItem.Text) Then
            List1.RemoveItem i
        End If
    Next
End Sub

Private Sub mnuVIG_Click()
    Dim i As Integer
    For i = 1 To LVU.ListItems.Count
        LVU.ListItems.Item(i).SmallIcon = Int(LVU.ListItems.Item(i).Tag)
    Next
    List1.Clear
End Sub

Private Sub mnuWhite_Click()
    Dim X As New frmWhiteBoard
    X.Show
End Sub

Private Sub Script_Error()
    '
End Sub

Private Sub Script_Timeout()
    '
End Sub

Private Sub Sock_Close()
    mnuConnect.Caption = "Connect"
    Sock.Close
    Unload frmLogin
    TBM.Buttons.Item(1).Enabled = True
    TBM.Buttons.Item(2).Enabled = False
    S.Panels.Item(1).Text = "Offline..."
    LVU.ListItems.Clear
    AddLog "<font Color=#990000>Disconnected from Server...</font>"
End Sub

Private Sub Sock_Connect()
    S.Panels.Item(1).Text = "Logging In..."
    AddLog "<font Color=#000099>Connected to Server...</font>"
    frmLogin.Show
    frmLogin.SetFocus
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
    'Recieve Msg From Client
    On Error Resume Next
    Dim pmm As Integer
    Dim Sata As String
    If Sock.State = 7 Then Sock.GetData Sata
    'MsgBox Sata
    DoEvents
    'Split Data
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
            Case "auth" '# Authentication error during create account or login
                If UBound(Stage2) < 1 Then GoTo err
                'Show Error
                Select Case Stage2(1)
                    Case "1"
                        MsgBox "Create account packet error!"
                    Case "2"
                        MsgBox "Create account invalid name/pssword size!"
                    Case "3"
                        MsgBox "Create account innalid charicters!"
                    Case "4"
                        MsgBox "Create account name in use!"
                    Case "5"
                        MsgBox "Login Packet error!"
                    Case "6"
                        MsgBox "Login invalid name/password size!"
                    Case "7"
                        MsgBox "Login invalid chrs!"
                    Case "8"
                        MsgBox "Login account not found!"
                    Case "9"
                        MsgBox "Login account in use!"
                    Case "10"
                        MsgBox "Unknown Login Err!"
                    Case "11"
                        MsgBox "Unknown Create Err!"
                    Case "15"
                        MsgBox "Mail: Unable to send user not found."
                    Case "16"
                        MsgBox "Mail: Please wait 45 seconds before sending another mail message."
                        Exit Sub
                    Case Else
                        MsgBox """" & Stage2(1) & """"
                End Select
            frmLogin.Show
            Case "topic"
                If UBound(Stage2) < 1 Then GoTo err
                AddLog "<font color=""#FF9900""> Topic : " & Stage2(1) & "</font>"
                S.Panels.Item(2).Text = Stage2(1)
            Case "list", "join" 'name,status
                S.Panels.Item(1).Text = "Online..."
                If UBound(Stage2) < 1 Then
                    GoTo err 'Require Variables
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 1 Then 'Require 2 Variables
                    GoTo err
                End If
                LVU.ListItems.Add , LCase(Stage3(0)), Stage3(0), , Int(Stage3(1))
                If IsIgnored(Stage3(0)) = True Then
                    LVU.ListItems(LCase(Stage3(0))).SmallIcon = 8
                End If
                LVU.ListItems(LCase(Stage3(0))).Tag = Stage3(1)
                LVU.SortKey = 1
                LVU.Sorted = True
                LVU.Sorted = False
                If LCase(Stage2(0)) = "join" Then
                    If mnuScript.Checked = True Then Script.Run "Join", Stage3(0)
                    AddLog "<Font Color=#009900>" & Stage3(0) & " has joined...</font>"
                    If LCase(UserName) <> LCase(Stage3(0)) And mnuPops.Checked = True Then DisplayAlert Stage3(0) & " Has Joined...", 3000, 101
                End If
            Case "tm", "pm" 'User 'Typing msg
                If UBound(Stage2) < 1 Then GoTo err
                LVU.ListItems(LCase(Stage2(1))).Bold = True
                If LCase(Stage2(0)) = "tm" Then
                    LVU.ListItems(LCase(Stage2(1))).ForeColor = vbRed
                Else
                    LVU.ListItems(LCase(Stage2(1))).ForeColor = vbBlue
                    pmm = PMName(Stage2(1))
                    PRoom(pmm).S.SimpleText = Stage2(1) & " is typing..."
                    'set pm win status typing
                End If
            Case "st", "ps" 'User 'Stoped
                If UBound(Stage2) < 1 Then GoTo err
                LVU.ListItems(LCase(Stage2(1))).Bold = False
                LVU.ListItems(LCase(Stage2(1))).ForeColor = vbBlack
                If LCase(Stage2(0)) = "ps" Then
                    'set pm win status not typing
                    pmm = PMName(Stage2(1))
                    PRoom(pmm).S.SimpleText = "Last message received " & PRoom(pmm).LMsg & "..."
                End If
            Case "part" 'name
                If UBound(Stage2) < 1 Then GoTo err
                LVU.ListItems.Remove LVU.ListItems(LCase(Stage2(1))).Index
                AddLog "<Font Color=#990000>" & Stage2(1) & " has left...</font>"
            Case "msg", "pmsg"
            'Font|Size|Bold|Italic|Underline|Color|Msg
                If UBound(Stage2) < 1 Then
                    GoTo err 'Require Variables
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 7 Then 'Require 8 Variables
                    GoTo err
                End If
                Dim WriteMsg As String
                If IsIgnored(Stage3(7)) Then Exit Sub
                WriteMsg = "<font face=""" & Stage3(0) & """ size=""" & Int(Stage3(1) / 3) & """ color=""" & Stage3(5) & """>"
                If Stage3(2) = "True" Then WriteMsg = WriteMsg & " <b>"
                If Stage3(3) = "True" Then WriteMsg = WriteMsg & " <i>"
                If Stage3(4) = "True" Then WriteMsg = WriteMsg & " <u>"
                Dim Link() As String
                Dim Links() As String
                Link = Split(Stage3(6), "{{")
                If UBound(Link) > 0 Then
                    For z = 1 To UBound(Link)
                        Links = Split(Link(z), "}}")
                        If UBound(Links) > 0 Then
                            If InStr(Links(0), "://") = 0 Then Links(0) = "http://" & Links(0)
                            Link(z) = "<a href=""" & Replace(Links(0), ":", Chr(1)) & """ target=""_blank"">" & Replace(Links(0), ":", Chr(1)) & "</a>"
                        End If
                        Link(z) = Link(z) & " " & Links(1)
                        For k = 2 To UBound(Links)
                            Link(z) = Link(z) & "}}" & Links(k)
                        Next
                    Next
                End If
                Stage3(6) = ""
                For z = 0 To UBound(Link)
                    Stage3(6) = Stage3(6) & Link(z)
                Next
                Dim FEmote() As String
                FEmote = Split(EChar, ",")
                For z = 0 To UBound(FEmote)
                    Stage3(6) = Replace(Stage3(6), FEmote(z), EmoteCode(FEmote(z)), , , vbTextCompare)
                Next
                WriteMsg = WriteMsg & Stage3(6) & "</u></i></b></font>"
                WriteMsg = Replace(WriteMsg, Chr(1), ":")
                WriteMsg = "<font color=""#000099"">&lt;" & Stage3(7) & "&gt;</font> " & WriteMsg
                Dim Param(1) As String
                Param(0) = Stage3(7)
                Param(1) = WriteMsg
                If LCase(Stage2(0)) = "msg" Then
                    AddLog WriteMsg
                    If Me.WindowState = vbMinimized And mnuSound.Checked = True Then PlaySound 102
                    If mnuScript.Checked = True Then Script.Run "MMsg", Stage3(7), Stage3(6)
                Else
                Dim Capname As String
                    If UBound(Stage3) < 8 Then
                        If UBound(Stage3) < 7 Then
                           GoTo err
                        End If
                        Capname = Stage3(7)
                        If mnuScript.Checked = True Then Script.Run "PMSG", Stage3(7), Stage3(6)
                    Else
                        Capname = Stage3(8)
                    End If
                        pmm = PMName(Capname)
                        If pmm > 0 Then
                            PRoom(pmm).LMsg = Now
                            PRoom(pmm).AddLog WriteMsg
                        Else
                            pmm = NewRoom
                            PRoom(pmm).Show
                            PRoom(pmm).Pnum = pmm
                            PRoom(pmm).Caption = Capname
                            PRoom(pmm).LMsg = Now
                            Pavail(pmm) = False
                            PRoom(pmm).AddLog WriteMsg
                        End If
                End If
            Case "server"
            If UBound(Stage2) < 1 Then GoTo err
            AddLog "<font color=""#3300FF"">Server : " & Stage2(1)
            Case "status"
                If UBound(Stage2) < 1 Then GoTo err
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 1 Then 'Require 2 Variables
                    GoTo err
                End If
                If IsIgnored(Stage3(0)) = False Then
                    LVU.ListItems(LCase(Stage3(0))).SmallIcon = Int(Stage3(1))
                Else
                    LVU.ListItems(LCase(Stage3(0))).SmallIcon = 8
                End If
                LVU.ListItems(LCase(Stage3(0))).Tag = Stage3(1)
            Case "inbox"
                If UBound(Stage2) < 1 Then
                    GoTo err 'Require Variables
                End If
                AddLog "<font color=""#FF9900""> Mail : " & Stage2(1) & " stored messages.</font>"
                S.Panels.Item(3).Text = "Inbox : " & Stage2(1)
            Case "sent"
                MsgBox "Mail Sent!"
                Exit Sub
            Case "newmail"
                If UBound(Stage2) < 1 Then
                    GoTo err 'Require Variables
                End If
                AddLog "<font color=""#FF9900""> Mail : New Mail From " & Stage2(1) & "...</font>"
                If mnuPops.Checked = True Then DisplayAlert " New Mail From " & Stage2(1) & "...", 3000, 103
            Case "endmail"
                If UBound(Stage2) < 1 Then
                    Exit Sub 'Require Variables
                End If
                MsgBox "Finished Downloading " & Stage2(1) & " stored message(s)."
                Exit Sub
            Case "mail"
                If UBound(Stage2) < 1 Then
                    Exit Sub 'Require Variables
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                If UBound(Stage3) < 4 Then 'Require 5 Variables
                    Exit Sub
                End If
                frmMail.LVM.ListItems.Add , Stage3(2), Stage3(2)
                frmMail.LVM.ListItems.Item(frmMail.LVM.ListItems.Count).Tag = Stage3(4)
                frmMail.LVM.ListItems.Item(frmMail.LVM.ListItems.Count).SubItems(1) = Stage3(1)
                frmMail.LVM.ListItems.Item(frmMail.LVM.ListItems.Count).SubItems(2) = Stage3(3)
                frmMail.LVM.ListItems.Item(frmMail.LVM.ListItems.Count).SubItems(3) = Stage3(0)
            Case "profile"
                If UBound(Stage2) < 1 Then
                    GoTo err 'Require Variables
                End If
                Stage3 = Split(Stage2(1), Chr(2))
                'If UBound(Stage3) < 7 Then 'Require 8 Variables
                '    GoTo err
                'End If
                Dim pfrm As New frmProfile
                pfrm.Show
                pfrm.Label1.Caption = Stage3(0)
                Dim fn
                For fn = 3 To 8
                    pfrm.txtLogin(fn).Text = Stage3(fn - 2)
                Next
                'If its your own Profile the server will let u edit it.
                If LCase(UserName) = LCase(Stage3(0)) Then
                    pfrm.Height = 4200
                End If
        End Select
err:
    Next
    
End Sub

Function IsIgnored(User As String) As Boolean
    Dim i As Integer
    IsIgnored = False
    For i = 0 To List1.ListCount - 1
        If LCase(List1.List(i)) = LCase(User) Then
            IsIgnored = True
            Exit For
        End If
    Next
End Function

Function EmoteCode(EPath As String)
    Dim Emote As String
    Select Case LCase(EPath)
        Case ":up"
        Emote = "icon_arrowu.gif"
        Case ":down"
        Emote = "icon_arrowd.gif"
        Case ":left"
        Emote = "icon_arrowl.gif"
        Case ":right"
        Emote = "icon_arrow.gif"
        Case ":d" ' big smile"
        Emote = "icon_biggrin.gif"
        Case ";d" ' cheesy grin
        Emote = "icon_cheesygrin.gif"
        Case ":s" 'confused
        Emote = "icon_confused.gif"
        Case ":c" ' cry
        Emote = "icon_cry.gif"
        Case ":cool"
        Emote = "icon_cool.gif"
        Case ":wow"
        Emote = "icon_eek.gif"
        Case ":6"
        Emote = "icon_evil.gif"
        Case ":!"
        Emote = "icon_exclaim.gif"
        Case ":(" ' frown"
        Emote = "icon_frown.gif"
        Case ":i"
        Emote = "icon_idea.gif"
        Case ":lol"
        Emote = "icon_lol.gif"
        Case ":m" ' Mad
        Emote = "icon_mad.gif"
        Case ":g" 'Green
        Emote = "icon_mrgreen.gif"
        Case ":|" 'nutral
        Emote = "icon_neutral.gif"
        Case ":?" ' Question
        Emote = "icon_question.gif"
        Case ":h" 'razz
        Emote = "icon_razz.gif"
        Case ":%" ' redface
        Emote = "icon_redface.gif"
        Case ":e" ' rolleys
        Emote = "icon_rolleyes.gif"
        Case ":[" ' sad
        Emote = "icon_sad.gif"
        Case ":)" ' Smile
        Emote = "icon_smile.gif"
        Case ":o" ' suprized
        Emote = "icon_surprised.gif"
        Case ":t" 'Twisted
        Emote = "icon_twisted.gif"
        Case ";)" 'Wink
        Emote = "icon_wink.gif"
    End Select
    If Len(Emote) > 0 Then
        EmoteCode = "<img src=""" & App.Path & "\Emoticons\" & Emote & """ />"
    Else
        EmoteCode = EPath
    End If
End Function

Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Sock_Close
End Sub

Private Sub TBF_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error:
    Select Case Button.Index
        Case 1
            CD.Flags = 1
            CD.FontName = txtChat.FontName
            CD.FontBold = txtChat.FontBold
            CD.FontItalic = txtChat.FontItalic
            CD.FontSize = txtChat.FontSize
            CD.FontUnderline = txtChat.FontUnderline
            CD.ShowFont
            If Len(CD.FontName) = 0 Then GoTo Error:
            If CD.FontSize > 14 Then CD.FontSize = 14 'Server Wont Allows any Bigger
            txtChat.FontName = CD.FontName
            txtChat.FontBold = CD.FontBold
            txtChat.FontItalic = CD.FontItalic
            txtChat.FontSize = CD.FontSize
            txtChat.FontUnderline = CD.FontUnderline
        Case 3
            If TBF.Buttons.Item(3).MixedState = True Then
                txtChat.FontBold = False
            Else
                txtChat.FontBold = True
            End If
        Case 4
            If TBF.Buttons.Item(4).MixedState = True Then
                txtChat.FontItalic = False
            Else
                txtChat.FontItalic = True
            End If
        Case 5
            If TBF.Buttons.Item(5).MixedState = True Then
                txtChat.FontUnderline = False
            Else
                txtChat.FontUnderline = True
            End If
        Case 7
            CD.Color = txtChat.ForeColor
            CD.ShowColor
            txtChat.ForeColor = CD.Color
        Case 9
            Dim GURL As String
            GURL = InputBox("What link would you like to add?", "URL", "http://www.pscode.com")
            If Len(GURL) > 0 Then txtChat.Text = txtChat.Text & "{{" & GURL & "}}"
        Case 10
            'Shell "explorer.exe """ & App.Path & "\Emoticons\index.html" & """", vbMaximizedFocus
            frmIcons.Show
        Case 12
            frmBG.Show
    
    End Select
Error:
    If txtChat.FontUnderline = True Then
        TBF.Buttons.Item(5).MixedState = True
    Else
        TBF.Buttons.Item(5).MixedState = False
    End If
    If txtChat.FontBold = True Then
        TBF.Buttons.Item(3).MixedState = True
    Else
        TBF.Buttons.Item(3).MixedState = False
    End If
    If txtChat.FontItalic = True Then
        TBF.Buttons.Item(4).MixedState = True
    Else
        TBF.Buttons.Item(4).MixedState = False
    End If
End Sub

Private Sub TBM_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1, 2
            mnuConnect_Click
        Case 4
            SaveHtml Me.Web
        Case 5
            ClearLog
        Case 7
            frmScript.Show
    End Select
End Sub

Public Function SockSend(Data As String)
'msgMS Sans Serif8.25FalseFalseFalse#000000
    If Sock.State = 7 Then
        Sock.SendData Data & Chr(0)
    End If
End Function

Private Sub TBM_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    SockSend "status" & Chr(1) & ButtonMenu.Index
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Web.Document.body.scrolltop = CLng(Len(Web.Document.body.innerHTML)) * 1000
Timer1.Enabled = False
End Sub

Private Sub tmrTimeout_Timer()
    SockSend ""
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    SockSend Chr(5)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdSend_Click
    End If
End Sub

Public Function AddLog(msg As String)
    On Error Resume Next
    Web.Document.body.innerHTML = Web.Document.body.innerHTML & msg & "<br>"
    Web.Document.body.scrolltop = CLng(Len(Web.Document.body.innerHTML)) * 1000
End Function

Function ClearLog()
'On Error Resume Next
    Web.Navigate "about:blank"
    Do While Web.ReadyState <> READYSTATE_COMPLETE
      DoEvents
    Loop
    DoEvents
    Web.Document.body.Style.WordWrap = "break-word"
    Web.Document.body.topmargin = 5
    Web.Document.body.leftmargin = 3
    Web.Document.body.innerHTML = "<b><i><font color=""#999999"">PSC Chat V5 Build 37</font></b></i><br>"
    Web.Document.body.Style.background = GetSetting("PSC5", "Options", "BG", "url(" & App.Path & "\" & "Top Right.jpg) #ffffff fixed no-repeat right top")
End Function

Private Sub Web_GotFocus()
    txtChat.SetFocus
End Sub

