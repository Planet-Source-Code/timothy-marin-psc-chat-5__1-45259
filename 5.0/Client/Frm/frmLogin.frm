VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   2895
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmNewUser 
      Height          =   3255
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Width           =   2895
      Begin VB.TextBox txtLogin 
         Height          =   285
         Index           =   8
         Left            =   1035
         TabIndex        =   10
         Top             =   2400
         Width           =   1710
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Index           =   7
         Left            =   1035
         TabIndex        =   9
         Top             =   2040
         Width           =   1710
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Create"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Index           =   6
         Left            =   1035
         TabIndex        =   8
         Top             =   1680
         Width           =   1710
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Index           =   5
         Left            =   1035
         TabIndex        =   7
         Top             =   1320
         Width           =   1710
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Index           =   4
         Left            =   1035
         TabIndex        =   6
         Top             =   960
         Width           =   1710
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Index           =   3
         Left            =   1035
         TabIndex        =   5
         Top             =   600
         Width           =   1710
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Index           =   2
         Left            =   1035
         TabIndex        =   4
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Age :"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   22
         Top             =   2445
         Width           =   855
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Country :"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   21
         Top             =   2085
         Width           =   855
      End
      Begin VB.Label lblCreate 
         Caption         =   "-Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Photo(Url) :"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   19
         Top             =   1725
         Width           =   855
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Skills :"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   18
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Website :"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   17
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   16
         Top             =   645
         Width           =   855
      End
      Begin VB.Label lblLogin 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retype :"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   120
      Picture         =   "frmLogin.frx":038A
      ScaleHeight     =   2715
      ScaleWidth      =   2595
      TabIndex        =   23
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   1005
      Width           =   975
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Index           =   1
      Left            =   1050
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   1
      Top             =   195
      Width           =   1695
   End
   Begin VB.Label lblCreate 
      BackStyle       =   0  'Transparent
      Caption         =   "+New Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1035
      Width           =   1335
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   645
      Width           =   855
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLogin_Click(Index As Integer)
UserName = txtLogin(0)
Password = txtLogin(1)
SaveSetting "PSC5", "Options", "Name", UserName
SaveSetting "PSC5", "Options", "Password", Password
If Index = 1 Then
    For i = 0 To txtLogin.UBound
        If txtLogin(i) = "" Then
            MsgBox "Please fill out all feilds. You may use N/A"
            Exit Sub
        End If
        If IChr(txtLogin(i).Text) <> txtLogin(i).Text Then
            MsgBox "One or more field(s) contains invalid charicters. Please use alphanumeric and punctuation exept <>."
            Exit Sub
        End If
    Next
    If txtLogin(1) <> txtLogin(2) Then
        MsgBox "Password does not match Retype"
        Exit Sub
    End If
    Dim Send As String
    Send = "create" & Chr(1) & txtLogin(0) & Chr(2) & txtLogin(1)
    For i = 3 To txtLogin.UBound
        Send = Send & Chr(2) & txtLogin(i)
    Next
    frmMain.SockSend Send
Else
        For i = 0 To 1
        If txtLogin(i) = "" Then
            MsgBox "Please fill out all feilds. You may use N/A"
            Exit Sub
        End If
        If IChr(txtLogin(i).Text) <> txtLogin(i).Text Then
            MsgBox "One or more field(s) contains invalid charicters. Please use alphanumeric and punctuation exept <>."
            Exit Sub
        End If
    Next
    Send = "login" & Chr(1) & txtLogin(0) & Chr(2) & txtLogin(1)
    frmMain.SockSend Send
End If
    Unload Me
End Sub

Private Sub Form_Load()
    frmNewUser.Visible = False
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    txtLogin(0).Text = GetSetting("PSC5", "Options", "Name", "")
    txtLogin(1).Text = GetSetting("PSC5", "Options", "Password", "")
End Sub

Private Sub lblCreate_Click(Index As Integer)
    If Index = 0 Then
        frmNewUser.Visible = True
        cmdLogin(0).Visible = False
    Else
        frmNewUser.Visible = False
        cmdLogin(0).Visible = True
    End If
End Sub
