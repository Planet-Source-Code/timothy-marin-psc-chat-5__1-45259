VERSION 5.00
Begin VB.Form frmProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Apply Changes"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Index           =   3
      Left            =   1275
      TabIndex        =   1
      Top             =   840
      Width           =   3270
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Index           =   4
      Left            =   1275
      TabIndex        =   2
      Top             =   1200
      Width           =   3270
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Index           =   5
      Left            =   1275
      TabIndex        =   3
      Top             =   1560
      Width           =   3270
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Index           =   6
      Left            =   1275
      TabIndex        =   4
      Top             =   1920
      Width           =   3270
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Index           =   7
      Left            =   1275
      TabIndex        =   5
      Top             =   2280
      Width           =   3270
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Index           =   8
      Left            =   1275
      TabIndex        =   6
      Top             =   2640
      Width           =   3270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   4800
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   4680
      Y1              =   3135
      Y2              =   3120
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   885
      Width           =   855
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Website :"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   1245
      Width           =   855
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Skills :"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   1605
      Width           =   855
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Photo(Url) :"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   1965
      Width           =   855
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Country :"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Top             =   2325
      Width           =   855
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Age :"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   2685
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    MsgBox "Please note this will only modify the profile of the name you are logged in as. To see if the profile change was succesfull check your profile again."
    Dim snd As String
    snd = Me.txtLogin(3).Text
    For i = 4 To 8
        snd = snd & Chr(2) & txtLogin(i).Text
    Next
    frmMain.SockSend "ep" & Chr(1) & snd
    Unload Me
End Sub
