VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "AOL Winsock Example by GodsMisfit"
   ClientHeight    =   4665
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4425
   Height          =   5070
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   4425
   Top             =   1170
   Width           =   4545
   Begin VB.CommandButton sendim 
      Caption         =   "Send IM"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox immsg 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox imname 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton connect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox password 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox screenname 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox status 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
   Begin TTOSocket.TTOSock TTOSock1 
      Left            =   1560
      Top             =   2040
      _ExtentX        =   979
      _ExtentY        =   953
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim sixth As String
Dim seventh As String
Dim packet As String
Dim aolsock As Long
Private Sub connect_Click()
TTOSock1.ConnectTo "AmericaOnline.aol.com", 5190 'This opens a connection to the AOL server


End Sub


Private Sub sendim_Click()
screenname1$ = Int2Hex(Len(imname.Text)) & EnHex(imname.Text)
IM1$ = Int2Hex(Len(immsg.Text)) & EnHex(immsg.Text)
If sixth = "127" Then
sixth = "15" 'Backtrack 1 because we are going to add 1 soon
seventh = "23" 'THIS IS THE CODE TO RESET THE BYTES
End If
sixth = CInt(sixth) + 1
seventh = CInt(seventh) + 1
fifth$ = give5th("5A606700001522A06953002500010001070400000003010A04000000010301" & screenname1$ & "011D00010A04000000020301" & IM1$ & "011D00011D00011D000002000D", 2)
packet$ = DeHex("00" & fifth$ & Int2Hex(sixth) & Int2Hex(seventh) & "A06953002500010001070400000003010A04000000010301" & screenname1$ & "011D00010A04000000020301" & IM1$ & "011D00011D00011D000002000D")
crc1 = AOLCRC(packet$, Len(packet$) - 1)
thehi$ = Chr(gethibyte(crc1))
thelo$ = Chr(getlobyte(crc1))
packet$ = DeHex("5A" & EnHex(thehi$) & EnHex(thelo$) & "00" & fifth$ & Int2Hex(sixth) & Int2Hex(seventh) & "A06953002500010001070400000003010A04000000010301" & screenname1$ & "011D00010A04000000020301" & IM1$ & "011D00011D00011D000002000D")
TTOSock1.SendDataTo packet$, aolsock
immsg.Text = ""

End Sub

Private Sub status_Change()
status.SelStart = Len(status.Text) 'Auto Scrolls
If Len(status.Text) >= 32000 Then status.Text = Right$(status.Text, Len(status.Text) / 2) 'Keeps the length from getting too large
End Sub


Private Sub TTOSock1_Connected(ByVal SocketID As Long)
aolsock = SocketID
status.Text = status.Text & "Connected" & vbCrLf
packet$ = DeHex("5A413800347F7FA3036B0100F5000000050F00002152CBCA070A1000080400000000035F0000010004000300080000000000000000000000020D")
TTOSock1.SendDataTo packet$, aolsock ' This sends the version packet, NO CHANGES HAVE TO BE MADE

End Sub

Private Sub TTOSock1_DataArrival(ByVal SocketID As Long, sData As String)
If Len(sData) = 9 Then
ping6$ = Mid(sData, 6, 1)
ping7$ = Mid(sData, 7, 1)
packet$ = DeHex("0003" & EnHex(ping7$) & EnHex(ping6$) & "A40D")
crc1 = AOLCRC(packet$, Len(packet$) - 1)
thehi$ = Chr(gethibyte(crc1))
thelo$ = Chr(getlobyte(crc1))
tosend$ = DeHex("5A" & EnHex(thehi$) & EnHex(thelo$) & EnHex(packet$))
TTOSock1.SendDataTo tosend$, aolsock
End If

Dim crap As Integer
crap = 1
Do Until crap = 0
crap = InStr(1, sData, Chr$(0))
If crap > 0 Then Mid(sData, crap, 1) = "Å"
Loop
status.Text = status.Text & "Server:  " & sData & vbCrLf & vbCrLf

If InStr(1, sData, "Invalid password") Then
status.Text = status.Text & "Invalid Password" & vbCrLf
TTOSock1.Disconnect aolsock
End If
If InStr(1, sData, "Invalid account") Then
status.Text = status.Text & "Invalid account" & vbCrLf
TTOSock1.Disconnect aolsock
End If

If InStr(1, sData, "SD") Then
screenname1$ = Int2Hex(Len(screenname.Text) + 1) & EnHex(screenname.Text)
password1$ = Int2Hex(Len(password.Text)) & EnHex(password.Text)
fifth$ = give5th("5AD69400411010A044640015000100010A0400000001010B04000000010301" & screenname1$ & "20011D00011D00010A04000000020301" & password1$ & "011D000002000D", 2)
packet$ = DeHex("00" & fifth$ & "1010A044640015000100010A0400000001010B04000000010301" & screenname1$ & "20011D00011D00010A04000000020301" & password1$ & "011D000002000D")
crc1 = AOLCRC(packet$, Len(packet$) - 1)
thehi$ = Chr(gethibyte(crc1))
thelo$ = Chr(getlobyte(crc1))
packet$ = DeHex("5A" & EnHex(thehi$) & EnHex(thelo$) & "00" & fifth$ & "1010A044640015000100010A0400000001010B04000000010301" & screenname1$ & "20011D00011D00010A04000000020301" & password1$ & "011D000002000D")
TTOSock1.SendDataTo packet$, aolsock 'Sends the screenname/password.  If you don’t understand this refer above to 'Forming Packets'
packet$ = DeHex("5A3A0A00031018A40D") 'This is a logged ping packet
TTOSock1.SendDataTo packet$, aolsock
End If

If InStr(1, sData, "Master Tool") Then
packet$ = DeHex("5A3A0A00031018A40D5AC6910008111CA079610701010D5A5C340022121CA0746C00170001000A0F04007BB5E80A10041DA6DDF20A4504000000070002000D5A6F4A000D131CA0534300150001000002000D")
TTOSock1.SendDataTo packet$, aolsock 'This sends the ya login packet, NO CHANGES HAVE BEEN MADE
sixth = Hex2Int("13")
seventh = Hex2Int("21") 'This is where you set your sixth and seventh at first.
End If

End Sub


Private Sub TTOSock1_SendComplete(ByVal SocketID As Long)
Dim crap As Integer
crap = 1
Do Until crap = 0
crap = InStr(1, packet$, Chr$(0))
If crap > 0 Then Mid(packet$, crap, 1) = "Å"
Loop
status.Text = status.Text & "Client:  " & packet$ & vbCrLf & vbCrLf

End Sub


