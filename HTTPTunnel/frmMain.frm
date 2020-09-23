VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "HTTP Tunnel/Bouncer"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmiplog 
      BackColor       =   &H00000000&
      Caption         =   "IP Log"
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   7080
      TabIndex        =   11
      Top             =   120
      Width           =   2535
      Begin VB.ListBox lstips 
         Height          =   2010
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame frmData 
      BackColor       =   &H00000000&
      Caption         =   "Data"
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtlog 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame frmConnection 
      BackColor       =   &H00000000&
      Caption         =   "Connection"
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdListen 
         Caption         =   "Activate Tunnel"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtport 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txthostip 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtlocalport 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblstatus 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tunnel Down"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lbl 
         BackColor       =   &H00000000&
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lbl 
         BackColor       =   &H00000000&
         Caption         =   "Host Or IP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00000000&
         Caption         =   "Listen On Port:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Left            =   2160
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrsckserver 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2760
      Top             =   1800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdListen_Click()
sckTCP.Close
sckTCP.LocalPort = txtlocalport
sckTCP.Listen
lblstatus.Caption = "Tunnel Open"
lblstatus.ForeColor = &HFFFF00
End Sub

Private Sub Form_Load()
Call WSAStartup(&H101, WSAInfo)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lngSocketHandle <> 0 Then Call CloseSocket(lngSocketHandle)
Call WSACleanup
End Sub

Private Sub Form_Resize()
On Error Resume Next
frmMain.Width = 9870
frmMain.Height = 3180
End Sub

Private Sub sckTCP_Close()
lblstatus.Caption = "Tunnel on Standby"
lblstatus.ForeColor = &HFFFF&
sckTCP.Close
sckTCP.Listen
End Sub

Private Sub sckTCP_ConnectionRequest(ByVal requestID As Long)

lblstatus.Caption = "Tunnel Active"
lblstatus.ForeColor = &HFF00&
    Call CloseSocket(lngSocketHandle)
    lngSocketHandle = 0
    Call sckserverconnect
    
    sckTCP.Close
    sckTCP.Accept requestID
    lstips.AddItem sckTCP.RemoteHostIP

End Sub

Private Sub sckTCP_DataArrival(ByVal bytesTotal As Long)
Dim temp As String
If sckTCP.State = sckConnected Then sckTCP.GetData temp
txtlog.Text = txtlog & temp & vbCrLf
lblstatus.Caption = "Recieving"
lblstatus.ForeColor = &H80FF&
Call sckserversenddata(temp)
End Sub

Private Sub txtlog_Change()
txtlog.SelLength = Len(txtlog.Text)
If Len(txtlog) > 30000 Then
txtlog.Text = ""
End If
End Sub

Private Function sckserverconnect()

  Dim lngHostName As String
  Dim udtSockaddr As SOCKADDR_IN
    
    Call ioctlsocket(lngSocketHandle, FIONBIO, 1)
    
    lngSocketHandle = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)

    If lngSocketHandle <= 0 Then
        txtlog.Text = txtlog.Text & "*** ERROR: Could not create socket!" & vbCrLf
        txtlog.SelStart = Len(txtlog)
        Exit Function
    End If
    
    Dim hostent_addr As Long
    Dim hostip_addr As Long
    Dim host As HOSTENT
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    hostent_addr = gethostbyname(txthostip.Text)


    If hostent_addr = 0 Then
        MsgBox "Can't resolve name."
        Exit Function
    End If
    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
    ReDim temp_ip_address(1 To host.hLength)
    RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength


    For i = 1 To host.hLength
        ip_address = ip_address & temp_ip_address(i) & "."
    Next
    ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

    lngHostName = inet_addr(ip_address)
    
    With udtSockaddr
         .sin_family = AF_INET
         .sin_addr = lngHostName
         .sin_port = htons(txtport.Text)
    End With
    
    Call ioctlsocket(lngSocketHandle, FIONBIO, 0)
    
    If Connect(lngSocketHandle, udtSockaddr, Len(udtSockaddr)) = -1 Then
        txtlog.Text = txtlog.Text & "*** ERROR: Cannot Connect to " & txthostip.Text & vbCrLf
        txtlog.SelStart = Len(txtlog)
        Call CloseSocket(lngSocketHandle)
        Exit Function
    Else
        txtlog.Text = txtlog.Text & "*** Connected to " & txthostip.Text & " on " & txtport.Text & vbCrLf
        txtlog.SelStart = Len(txtlog)
    End If
    
    Call ioctlsocket(lngSocketHandle, FIONBIO, 1)
    tmrsckserver.Enabled = True

End Function

Sub sckserversenddata(buffer As String)
On Error Resume Next

  Dim arrBuffer()     As Byte
  Dim strdata         As String
  Dim BytesSent       As Long
      
        strdata = buffer

        arrBuffer() = StrConv(strdata, vbFromUnicode)
        
        BytesSent = Send(lngSocketHandle, arrBuffer(0), Len(strdata), 0&)
                
End Sub

Private Sub tmrsckserver_Timer()
On Error Resume Next
  Const MAX_BUFFER_LENGTH As Long = 8192

  Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
  Dim lngBytesReceived                    As Long
  Dim strTempBuffer                       As String
    
    Do
        lngBytesReceived = Recv(lngSocketHandle, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)
        DoEvents
        
        If lngBytesReceived > 0 Then
        
            strTempBuffer = StrConv(arrBuffer, vbUnicode)

            strTempBuffer = Left$(strTempBuffer, lngBytesReceived)
        
            txtlog.Text = txtlog & strTempBuffer & vbCrLf
lblstatus.Caption = "Sending"
lblstatus.ForeColor = &H80FF&
sckTCP.SendData strTempBuffer
        Else

            Exit Do
        End If
    Loop

End Sub
