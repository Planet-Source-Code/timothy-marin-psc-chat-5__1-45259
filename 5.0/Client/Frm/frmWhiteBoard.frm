VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWhiteBoard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WBoard"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "frmWhiteBoard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar S 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4710
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   3960
      Width           =   7335
      Begin VB.CommandButton Command1 
         Caption         =   "Listen"
         Height          =   300
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   300
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "15006"
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   4710
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   8308
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Color"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Line"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      Height          =   4575
      Left            =   720
      MousePointer    =   2  'Cross
      ScaleHeight     =   4515
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   6840
         Top             =   960
      End
      Begin MSWinsockLib.Winsock SckClient 
         Left            =   6840
         Top             =   2880
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock SckServer 
         Left            =   6840
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6720
         Top             =   360
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
               Picture         =   "frmWhiteBoard.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWhiteBoard.frx":0724
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWhiteBoard.frx":69BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWhiteBoard.frx":6D58
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Line LineR 
         Index           =   0
         Visible         =   0   'False
         X1              =   120
         X2              =   360
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         Index           =   0
         Visible         =   0   'False
         X1              =   120
         X2              =   360
         Y1              =   120
         Y2              =   120
      End
   End
End
Attribute VB_Name = "frmWhiteBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------
'Copyright (C) IntraDream, Inc 2002 - 2003
'If you have any questions Please Contact
'IntraDreams Support; Support@intradream.com
'
'This program is free software; you can redistribute
'it and/or modify it under the terms of the GNU General
'Public License as published by the Free Software
'Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will
'be useful, but WITHOUT ANY WARRANTY; without even the
'implied warranty of MERCHANTABILITY or FITNESS FOR A
'PARTICULAR PURPOSE. See the GNU General Public License
'for more details.
'
'You should have received a copy of the GNU General Public
'License along with this program; if not, write to the Free
'Software Foundation, Inc., 59 Temple Place, Suite 330,
'Boston, MA 02111-1307 USA
'----------------------------------

Dim ButtonPress As Boolean
Dim lastx, lasty As Integer
Dim lastxr, lastyr As Integer
Public intType As Integer '1 for Server 2 Client

Private Sub cmdConnect_Click()
    On Error Resume Next
    SckClient.Close
    SckServer.Close
    SckClient.Connect Text1.Text, txtPort.Text
    S.SimpleText = "connecting to " & Text1.Text & ":" & txtPort.Text & "..."
    Frame1.Visible = False
End Sub

Private Sub Command1_Click()
    On Error GoTo err
    SckClient.Close
    SckServer.Close
    SckServer.LocalPort = txtPort.Text
    SckServer.Listen
    Frame1.Visible = False
    S.SimpleText = "listening on " & txtPort.Text & "..."
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
Picture1.Enabled = True
Line1(0).BorderColor = vbBlack
Line1(0).BorderWidth = 3
    S.SimpleText = "disconnected..."
    SckClient.Close
    SckServer.Close
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ButtonPress = True
   On Error Resume Next
   If SckClient.State = 7 Then
      If SckClient.State <> 7 Then Exit Sub
      SckClient.SendData "New||" & X & "||" & Y & vbNewLine
      lastx = X
      lasty = Y
     Else
      If SckServer.State <> 7 Then Exit Sub
      SckServer.SendData "New||" & X & "||" & Y & vbNewLine
      lastx = X
      lasty = Y
    End If
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If ButtonPress = True Then
     If SckClient.State = 7 Then
        SckClient.SendData "Move||" & X & "||" & Y & vbNewLine
        Load Line1(Line1.Count)
        Line1(Line1.UBound).X1 = lastx
        Line1(Line1.UBound).X2 = X
        Line1(Line1.UBound).Y1 = lasty
        Line1(Line1.UBound).Y2 = Y
        Line1(Line1.UBound).Visible = True
      
        lastx = X
        lasty = Y
      Else
        If SckServer.State <> 7 Then Exit Sub
        SckServer.SendData "Move||" & X & "||" & Y & vbNewLine
        Load Line1(Line1.Count)
        Line1(Line1.UBound).X1 = lastx
        Line1(Line1.UBound).X2 = X
        Line1(Line1.UBound).Y1 = lasty
        Line1(Line1.UBound).Y2 = Y
        Line1(Line1.UBound).Visible = True
        
        lastx = X
        lasty = Y
     End If
   End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ButtonPress = False
End Sub

Private Sub SckClient_Close()
    S.SimpleText = "disconnected..."
    SckClient.Close
End Sub

Private Sub SckClient_Connect()
  'Sends Status About Color, size etc
  On Error Resume Next
  Picture1.Enabled = True
  Line1(0).BorderColor = vbBlack
  Line1(0).BorderWidth = 3
  S.SimpleText = "connected..."
  If SckClient.State = 7 Then
     SckClient.SendData "Color||" & Line1(0).BorderColor & vbNewLine & "Size||" & Line1(0).BorderWidth & vbNewLine & "Yours?||" & vbNewLine
  End If
  
     
End Sub

Private Sub SckClient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
   Dim var1 As Variant
   Dim Var2 As Variant
   Dim strIn As String
   Dim lngA As Long
   Dim lngB As Long
   
   SckClient.GetData strIn
   var1 = Split(strIn, vbNewLine)
   
   For lngA = 0 To UBound(var1) - 1
      Var2 = Split(var1(lngA), "||")
      
      Select Case Var2(0)
         Case "Color"
            LineR(0).BorderColor = Var2(1)
         Case "Size"
            LineR(0).BorderWidth = Var2(1)
         Case "Yours?"
            SckClient.SendData "Color||" & Line1(0).BorderColor & vbNewLine & "Size||" & Line1(0).BorderWidth & vbNewLine
         Case "New"
            lastxr = Var2(1)
            lastyr = Var2(2)
         Case "Move"
            Load LineR(LineR.Count)
            LineR(LineR.UBound).X1 = lastxr
            LineR(LineR.UBound).X2 = Var2(1)
            LineR(LineR.UBound).Y1 = lastyr
            LineR(LineR.UBound).Y2 = Var2(2)
            LineR(LineR.UBound).Visible = True
            lastxr = Var2(1)
            lastyr = Var2(2)
         Case "Clear"
            Dim i As Long
            For i = 1 To Line1.Count - 1
               Unload Line1(i)
            Next
                    
            For i = 1 To LineR.Count - 1
               Unload LineR(i)
            Next
      End Select
   Next lngA
   
End Sub

Private Sub SckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    S.SimpleText = "disconnected..."
    SckClient.Close
End Sub

Private Sub SckServer_Close()
    S.SimpleText = "disconnected..."
    SckServer.Close
End Sub

Private Sub SckServer_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
  'Connects Other user
  Picture1.Enabled = True
  Line1(0).BorderColor = vbBlack
  Line1(0).BorderWidth = 3
  SckServer.Close
  SckServer.Accept requestID
  S.SimpleText = "connected..."
End Sub

Private Sub SckServer_DataArrival(ByVal bytesTotal As Long)
   On Error Resume Next
   Dim var1 As Variant
   Dim Var2 As Variant
   Dim strIn As String
   Dim lngA As Long
   Dim lngB As Long
   
   SckServer.GetData strIn
   var1 = Split(strIn, vbNewLine)
   
   For lngA = 0 To UBound(var1) - 1
      Var2 = Split(var1(lngA), "||")
      
      Select Case Var2(0)
         Case "Color"
            LineR(0).BorderColor = Var2(1)
         Case "Size"
            LineR(0).BorderWidth = Var2(1)
         Case "Yours?"
            SckServer.SendData "Color||" & Line1(0).BorderColor & vbNewLine & "Size||" & Line1(0).BorderWidth & vbNewLine
         Case "New"
            lastxr = Var2(1)
            lastyr = Var2(2)
         Case "Move"
            Load LineR(LineR.Count)
            LineR(LineR.UBound).X1 = lastxr
            LineR(LineR.UBound).X2 = Var2(1)
            LineR(LineR.UBound).Y1 = lastyr
            LineR(LineR.UBound).Y2 = Var2(2)
            LineR(LineR.UBound).Visible = True
            lastxr = Var2(1)
            lastyr = Var2(2)
         Case "Clear"
            Dim i As Long
            For i = 1 To Line1.Count - 1
               Unload Line1(i)
            Next
                    
            For i = 1 To LineR.Count - 1
               Unload LineR(i)
            Next
      End Select
   Next lngA
   
End Sub

Private Sub SckServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    S.SimpleText = "disconnected..."
    SckServer.Close
End Sub

Private Sub Timer1_Timer()
  Select Case intType
    Case 1
      If SckServer.State = 7 Then
         Picture1.Enabled = True
         Timer1.Enabled = False
      End If
    Case 2
      If SckClient.State = 7 Then
         Picture1.Enabled = True
         Timer1.Enabled = False
      End If
  End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Select Case Button.Caption
  
    Case "Color"
          Picture1.Enabled = True
          cd1.ShowColor
          
          'Adds the Colour to the RTB
          If cd1.Color <> 0 Then
             Line1(0).BorderColor = cd1.Color
                    
             If SckClient.State = 7 Then
                If SckClient.State <> 7 Then Exit Sub
                SckClient.SendData "Color||" & cd1.Color & vbNewLine
              Else
                If SckServer.State <> 7 Then Exit Sub
                SckServer.SendData "Color||" & cd1.Color & vbNewLine
             End If
          End If
          
    Case "Line"
          Dim lngSize As Long
          lngSize = InputBox("Please Enter the Size (1-10)")
          Line1(0).BorderWidth = lngSize
          
          If SckClient.State = 7 Then
             If SckClient.State <> 7 Then Exit Sub
             SckClient.SendData "Size||" & lngSize & vbNewLine
           Else
             If SckServer.State <> 7 Then Exit Sub
             SckServer.SendData "Size||" & lngSize & vbNewLine
          End If
          
    Case "Clear"
          If SckClient.State = 7 Then
             If SckClient.State <> 7 Then Exit Sub
             SckClient.SendData "Clear||" & vbNewLine
           Else
             If SckServer.State <> 7 Then Exit Sub
             SckServer.SendData "Clear||" & vbNewLine
          End If
          
          Dim i As Long
          For i = 1 To Line1.Count - 1
            Unload Line1(i)
          Next
                    
          For i = 1 To LineR.Count - 1
            Unload LineR(i)
          Next
  End Select
  
End Sub
