VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrivate 
   Caption         =   "Private"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5520
   Icon            =   "frmPrivate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar TBF 
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons(3)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
      EndProperty
   End
   Begin MSComctlLib.StatusBar S 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3780
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmSend 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   4575
      Begin VB.TextBox txtChat 
         Height          =   615
         Left            =   75
         MaxLength       =   200
         TabIndex        =   3
         Top             =   150
         Width           =   3540
      End
      Begin VB.PictureBox p 
         Height          =   630
         Left            =   3705
         ScaleHeight     =   570
         ScaleWidth      =   735
         TabIndex        =   1
         Top             =   150
         Width           =   790
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            Height          =   570
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   3
      Left            =   4680
      Top             =   1320
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
            Picture         =   "frmPrivate.frx":038A
            Key             =   "Small Caps"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":049C
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":05AE
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":06C0
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":07D2
            Key             =   "Button"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":08E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":0C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":1018
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":13B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivate.frx":194C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSA 
         Caption         =   "Save As"
      End
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "Extra"
      Begin VB.Menu mnuSendFile 
         Caption         =   "Send File"
      End
      Begin VB.Menu mnuWhiteboard 
         Caption         =   "Whiteboard"
      End
   End
End
Attribute VB_Name = "frmPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Pnum As Integer
Public LMsg As String

Function AddLog(msg As String)
    On Error Resume Next
    Web.Document.body.innerHTML = Web.Document.body.innerHTML & msg & "<br>"
    Web.Document.body.scrolltop = CLng(Len(Web.Document.body.innerHTML)) * 100
End Function

Function ClearLog()
On Error Resume Next
    Web.Navigate "about:blank"
    Do While Web.ReadyState <> READYSTATE_COMPLETE
      DoEvents
    Loop
    DoEvents
    Web.Document.body.topmargin = 5
    Web.Document.body.leftmargin = 3
    Web.Document.body.innerHTML = "<b><i><font color=""#999999"">PSC Chat V5 Build 37</font></b></i><br>"
End Function

Private Sub cmdSend_Click()
    'Font|Size|Bold|Italic|Underline|Color|Msg
    txtChat.SetFocus
    If Len(txtChat) < 1 Or frmMain.Sock.State <> 7 Then Exit Sub
    'MsgBox DectoWebCol(txtChat.ForeColor)
    frmMain.SockSend "pmsg" & Chr(1) & txtChat.Font & Chr(2) & txtChat.FontSize & Chr(2) & txtChat.FontBold & Chr(2) & txtChat.FontItalic & Chr(2) & txtChat.FontUnderline & Chr(2) & DectoWebCol(txtChat.ForeColor) & Chr(2) & txtChat.Text & Chr(2) & Me.Caption
    txtChat.Text = ""
End Sub

Private Sub Form_Load()
    txtChat.ForeColor = vbBlack
    ClearLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    frmSend.Width = Me.ScaleWidth + 10
    p.Left = Me.ScaleWidth - (cmdSend.Width + 120)
    txtChat.Width = Me.ScaleWidth - (cmdSend.Width + 250)
    TBF.Width = Me.ScaleWidth
    Web.Width = TBF.Width
    Web.Height = Me.ScaleHeight - (frmSend.Height + S.Height) + 100 - TBF.Height
    frmSend.Top = Web.Height + TBF.Height - 100
    TBF.Top = frmSend.Top - TBF.Height + 100
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Pavail(Pnum) = True
End Sub

Private Sub mnuSA_Click()
    frmMain.SaveHtml Me.Web
End Sub

Private Sub mnuSendFile_Click()
    Dim d As New frmTransfer
    d.Show
End Sub

Private Sub mnuWhiteboard_Click()
    Dim X As New frmWhiteBoard
    X.Show
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
            Shell "explorer.exe """ & App.Path & "\Emoticons\index.html" & """", vbMaximizedFocus
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


Private Sub txtChat_KeyPress(KeyAscii As Integer)
    frmMain.SockSend Chr(7) & Chr(1) & Me.Caption
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdSend_Click
    End If
End Sub
