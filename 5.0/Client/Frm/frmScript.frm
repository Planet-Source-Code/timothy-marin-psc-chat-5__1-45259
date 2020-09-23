VERSION 5.00
Begin VB.Form frmScript 
   Caption         =   "Script"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6375
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuLoadD 
         Caption         =   "Load Default"
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "Code"
      Begin VB.Menu mnuRcode 
         Caption         =   "Run Code"
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    txtScript.Height = Me.ScaleHeight
    txtScript.Width = Me.ScaleWidth
End Sub

Private Sub mnuLoad_Click()
    On Error Resume Next
    frmMain.CD.Filter = "Script File(*.txt)|*.txt"
    frmMain.CD.ShowOpen
    If Len(frmMain.CD.Filename) > 0 Then
        Dim FF As Integer
        Dim LI As String
        FF = FreeFile
        Open frmMain.CD.Filename For Input As #FF
            Line Input #FF, LI
            txtScript.Text = LI
            Do Until EOF(FF)
                DoEvents
                Line Input #FF, LI
                txtScript.Text = txtScript.Text & vbCrLf & LI
            Loop
        Close #FF
    End If
End Sub

Private Sub mnuLoadD_Click()
    txtScript.Text = "Sub MMsg(User,Msg)" & vbCrLf & "End Sub" & vbCrLf & vbCrLf & "Sub Join(User)" & vbCrLf & "End Sub" & vbCrLf & vbCrLf & "Sub PMSG(User,Msg)" & vbCrLf & "End Sub"
End Sub

Private Sub mnuRcode_Click()
    On Error GoTo err
    frmMain.ResetScript
    frmMain.Script.AddCode txtScript.Text
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuSave_Click()
    Dim FP As String
    FP = InputBox("What Name!", "PSCript", "PSCript")
    If Len(FP) < 1 Then Exit Sub
    Dim FF As Integer
    FF = FreeFile
    Open App.Path & "\" & FP & ".txt" For Output As #FF
        Print #FF, txtScript.Text
    Close #FF
    MsgBox "Save Succesfull!"
End Sub
