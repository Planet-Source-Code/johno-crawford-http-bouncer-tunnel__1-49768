Attribute VB_Name = "Winsock2API"
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Public Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long

Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long

Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long

Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function CloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long

Public Declare Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long

Public Declare Function Recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long


Public Const SOCK_STREAM = 1
Public Const AF_INET = 2
Public Const IPPROTO_TCP = 6


Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129

Public Const SCK_VERSION1 = &H101
Public Const SCK_VERSION2 = &H202

Public Const OFFSET = 65536

Public Const FIONBIO = &H8004667E

Public Type WSAData
    WVersion        As Integer
    WHighVersion    As Integer
    szDescription   As String * WSADESCRIPTION_LEN
    szSystemStatus  As String * WSASYS_STATUS_LEN
    iMaxSockets     As Integer
    iMaxUdpDg       As Integer
    lpVendorInfo    As Long
End Type

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
    End Type

Public Type SOCKADDR_IN
    sin_family          As Integer
    sin_port            As Integer
    sin_addr            As Long
    sin_zero(1 To 8)    As Byte
End Type

Public Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Public WSAInfo         As WSAData
Public lngSocketHandle As Long
