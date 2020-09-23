VERSION 5.00
Begin VB.Form frmBG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Background Information"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frmBG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmBG.frx":0582
      Left            =   2040
      List            =   "frmBG.frx":058F
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmBG.frx":05A2
      Left            =   360
      List            =   "frmBG.frx":05B2
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   2040
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Not Tiled"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Fixed"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.TextBox txtImage 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "..."
      Height          =   300
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Image"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "#FFFFFF"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2115
      Width           =   1455
   End
End
Attribute VB_Name = "frmBG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click(Index As Integer)
    Combo1.Enabled = Check1(1).Value
    Combo2.Enabled = Check1(1).Value
    If Combo1.Enabled = True Then
        Combo1.ListIndex = 2
        Combo2.ListIndex = 2
    End If
End Sub

Private Sub cmdApply_Click()
        frmMain.Web.Document.body.Style.background = "url('" & txtImage.Text & "') " & DectoWebCol(Picture1.BackColor)
        frmMain.Web.Document.body.Style.background = frmMain.Web.Document.body.Style.background & " " & DectoWebCol(Picture1.BackColor)
        
        If Check1(0).Value = vbChecked Then
            frmMain.Web.Document.body.Style.background = frmMain.Web.Document.body.Style.background & " fixed"
        End If
        If Check1(1).Value = vbChecked Then
            frmMain.Web.Document.body.Style.background = frmMain.Web.Document.body.Style.background & " no-repeat"
        End If
        frmMain.Web.Document.body.Style.background = frmMain.Web.Document.body.Style.background & " " & Combo1.Text & " " & Combo2.Text
            SaveSetting "PSC5", "Options", "BG", frmMain.Web.Document.body.Style.background
            Unload Me
End Sub

Private Sub cmdSel_Click()
    On Error Resume Next
    frmMain.CD.Filter = "image|*.bmp;*.jpg;*.gif"
    frmMain.CD.ShowOpen
    If Len(frmMain.CD.Filename) > 0 Then
        txtImage.Text = frmMain.CD.Filename
    End If
End Sub

Private Sub Form_Load()
    Picture1.BackColor = vbWhite
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
End Sub

Private Sub Picture1_Click()
            frmMain.CD.Color = Picture1.BackColor
            frmMain.CD.ShowColor
            Picture1.BackColor = frmMain.CD.Color
End Sub
