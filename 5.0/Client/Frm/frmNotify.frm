VERSION 5.00
Begin VB.Form frmNotify 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   0
      Picture         =   "frmNotify.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.Label lblAlert 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   0
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' API Declarations
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
' Constants
Const SM_CXFULLSCREEN = 16   ' Width of window client area
Const SM_CYFULLSCREEN = 17   ' Height of window client area
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

' Declarations
Private fX As Long
Private fY As Long
Private lngScaleX As Long
Private lngScaleY As Long
Private AlertIndex As Long

Private Sub Label2_Click(Index As Integer)
If AlertCount = AlertIndex Then AlertCount = 0
Me.Visible = False
End Sub

Private Sub lblAlert_Click()
    ' When user clicked the alertbox
    frmMain.Show
    frmMain.WindowState = vbNormal
End Sub

Private Sub lblAlert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show as hyperlink
    If lblAlert.FontUnderline = False Then
        lblAlert.FontUnderline = True
        lblAlert.ForeColor = RGB(0, 0, 255)
    End If
    
GoTo blahh
blahh:
End Sub

Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show text
    If lblAlert.FontUnderline = True Then
        lblAlert.FontUnderline = False
        lblAlert.ForeColor = &H0
    End If
End Sub

Public Sub Display(MessageText As String, Duration As Long, Sound As Long)

    Dim wFlags As Long, X As Long

    ' Increase the alert count
    AlertCount = AlertCount + 1
    If AlertCount >= 5 Then AlertCount = 1
    AlertIndex = AlertCount

    ' Set the message
    lblAlert.Caption = MessageText
    
    ' Set the duration
    tmrAlert.Interval = Duration

    ' Get the system metrics we need
    fX = GetSystemMetrics(SM_CXFULLSCREEN)
    fY = GetSystemMetrics(SM_CYFULLSCREEN)
    lngScaleX = Me.Width - Me.ScaleWidth
    lngScaleY = Me.Height - Me.ScaleHeight
    
    ' Size the form
    Me.Height = 90
    Me.Width = picBackground.Width + lngScaleX
    Me.Left = fX * Screen.TwipsPerPixelX - Me.Width - 200
    Me.Top = (fY * Screen.TwipsPerPixelY) - ((picBackground.Height + lngScaleY) * (AlertCount - 1)) + 250
    Me.Show
    
    ' Play Sound
    PlaySound Sound
    ' Draw the gradient background
    If Duration = 0 Then Picture2.Visible = True
    ' Open the alert box
     Call tmOpen
 
        
End Sub

Private Sub tmOpen()
    Dim curHeight As Long
    Dim newHeight As Long
    Me.Height = picBackground.Height
    Me.Top = Me.Top - picBackground.Height
        tmrAlert.Enabled = True
End Sub

Private Sub tmrAlert_Timer()
    ' Alert was displayed, now close it
    tmrAlert.Enabled = False
    Call tmClose
End Sub

Private Sub tmClose()
        If AlertCount = AlertIndex Then AlertCount = 0
        Unload Me
End Sub

