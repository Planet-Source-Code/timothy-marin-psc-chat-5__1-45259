VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail"
   ClientHeight    =   4080
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1005
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Inbox"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Read"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Compse"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Check"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6240
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMail.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMail.frx":0724
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMail.frx":0ABE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMail.frx":0E58
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frmNav 
      Caption         =   "Inbox"
      Height          =   3495
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   6735
      Begin MSComctlLib.ListView LVM 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5530
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Subject"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Key"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame frmNav 
      Caption         =   "Read"
      Height          =   3495
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   6735
      Begin VB.CommandButton cmdReply 
         Caption         =   "Reply"
         Height          =   255
         Left            =   5160
         TabIndex        =   8
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtRSubject 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   5775
      End
      Begin VB.TextBox txtrBody 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1440
         Width           =   6495
      End
      Begin VB.Label lblMail 
         Caption         =   "Message :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblMail 
         Caption         =   "Subject :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblMail 
         Caption         =   "From :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame frmNav 
      Caption         =   "Compose"
      Height          =   3495
      Index           =   4
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   6735
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   255
         Left            =   5160
         TabIndex        =   4
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   5775
      End
      Begin VB.TextBox txtSbody 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1440
         Width           =   6495
      End
      Begin VB.Label lblMail 
         Caption         =   "Message :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblMail 
         Caption         =   "Subject :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblMail 
         Caption         =   "To :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Menu mnua 
      Caption         =   "a"
      Visible         =   0   'False
      Begin VB.Menu mnuRead 
         Caption         =   "Read"
      End
      Begin VB.Menu mnuReply 
         Caption         =   "Reply"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReply_Click()
    On Error Resume Next
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    txtTo.Text = txtFrom.Text
    txtSubject.Text = "RE :" & txtRSubject.Text
End Sub

Private Sub cmdSend_Click()
    frmMain.SockSend "smail" & Chr(1) & txtTo.Text & Chr(2) & txtSubject.Text & Chr(2) & txtSbody.Text
End Sub

Private Sub LVM_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub LVM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnua
End Sub

Private Sub mnuDelete_Click()
    On Error Resume Next
    frmMain.SockSend "dmail" & Chr(1) & LVM.SelectedItem.SubItems(3)
    LVM.ListItems.Remove LVM.SelectedItem.Index
End Sub

Private Sub mnuRead_Click()
    On Error Resume Next
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
    txtFrom.Text = LVM.SelectedItem
    txtRSubject.Text = LVM.SelectedItem.SubItems(2)
    txtrBody.Text = LVM.SelectedItem.Tag
End Sub

Private Sub mnuReply_Click()
    On Error Resume Next
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    txtTo.Text = LVM.SelectedItem
    txtSubject.Text = "RE :" & LVM.SelectedItem.SubItems(2)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        On Error Resume Next
        If Button.Index = 6 Then
            LVM.ListItems.Clear
            frmMain.SockSend "cmail" & Chr(1)
            'Check Mail
        End If
        If Button.Index > 4 Then Exit Sub
        For i = 0 To 4
            frmNav(i).Visible = False
        Next
        frmNav(Button.Index).Visible = True
End Sub

