Attribute VB_Name = "Winsock_API"
'As per the release notes module, thanks to
' ** Daniel - sigsegv@mail.utexas.edu **
'for the majority of code in this module.

Option Explicit

Public Const FD_SETSIZE = 64
Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type

Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Public Const IPPROTO_TCP = 6

Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public Const sockaddr_size = 16
Public Const hostent_size = 16
Public Const INADDR_NONE = &HFFFFFFFF
Public Const INADDR_ANY = &H0

Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1

Public Const SOCK_STREAM = 1

Public Const AF_INET = 2

'Windows Sockets definitions of regular Microsoft C error constants
Global Const WSAEINTR = 10004
Global Const WSAEBADF = 10009
Global Const WSAEACCES = 10013
Global Const WSAEFAULT = 10014
Global Const WSAEINVAL = 10022
Global Const WSAEMFILE = 10024
'Windows Sockets definitions of regular Berkeley error constants
Global Const WSAEWOULDBLOCK = 10035
Global Const WSAEINPROGRESS = 10036
Global Const WSAEALREADY = 10037
Global Const WSAENOTSOCK = 10038
Global Const WSAEDESTADDRREQ = 10039
Global Const WSAEMSGSIZE = 10040
Global Const WSAEPROTOTYPE = 10041
Global Const WSAENOPROTOOPT = 10042
Global Const WSAEPROTONOSUPPORT = 10043
Global Const WSAESOCKTNOSUPPORT = 10044
Global Const WSAEOPNOTSUPP = 10045
Global Const WSAEPFNOSUPPORT = 10046
Global Const WSAEAFNOSUPPORT = 10047
Global Const WSAEADDRINUSE = 10048
Global Const WSAEADDRNOTAVAIL = 10049
Global Const WSAENETDOWN = 10050
Global Const WSAENETUNREACH = 10051
Global Const WSAENETRESET = 10052
Global Const WSAECONNABORTED = 10053
Global Const WSAECONNRESET = 10054
Global Const WSAENOBUFS = 10055
Global Const WSAEISCONN = 10056
Global Const WSAENOTCONN = 10057
Global Const WSAESHUTDOWN = 10058
Global Const WSAETOOMANYREFS = 10059
Global Const WSAETIMEDOUT = 10060
Global Const WSAECONNREFUSED = 10061
Global Const WSAELOOP = 10062
Global Const WSAENAMETOOLONG = 10063
Global Const WSAEHOSTDOWN = 10064
Global Const WSAEHOSTUNREACH = 10065
Global Const WSAENOTEMPTY = 10066
Global Const WSAEPROCLIM = 10067
Global Const WSAEUSERS = 10068
Global Const WSAEDQUOT = 10069
Global Const WSAESTALE = 10070
Global Const WSAEREMOTE = 10071
'Extended Windows Sockets error constant definitions
Global Const WSASYSNOTREADY = 10091
Global Const WSAVERNOTSUPPORTED = 10092
Global Const WSANOTINITIALISED = 10093
Global Const WSAHOST_NOT_FOUND = 11001
Global Const WSATRY_AGAIN = 11002
Global Const WSANO_RECOVERY = 11003
Global Const WSANO_DATA = 11004
Global Const WSANO_ADDRESS = 11004

#If Win16 Then
'---Windows System functions
    Public Declare Function PostMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
    Public Declare Sub MemCopy Lib "Kernel" Alias "hmemcpy" (Dest As Any, Src As Any, ByVal cb&)
    Public Declare Function lstrlen Lib "Kernel" (ByVal lpString As Any) As Integer
'---async notification constants
    Public Const SOL_SOCKET = &HFFFF
    Public Const SO_LINGER = &H80
    Public Const FD_READ = &H1
    Public Const FD_WRITE = &H2
    Public Const FD_OOB = &H4
    Public Const FD_ACCEPT = &H8
    Public Const FD_CONNECT = &H10
    Public Const FD_CLOSE = &H20
'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "Winsock.dll" (ByVal S As Integer, addr As sockaddr, addrlen As Integer) As Integer
    Public Declare Function bind Lib "Winsock.dll" (ByVal S As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
    Public Declare Function closesocket Lib "Winsock.dll" (ByVal S As Integer) As Integer
    Public Declare Function Connect Lib "Winsock.dll" Alias "connect" (ByVal S As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
    Public Declare Function ioctlsocket Lib "Winsock.dll" (ByVal S As Integer, ByVal cmd As Long, argp As Long) As Integer
    Public Declare Function getpeername Lib "Winsock.dll" (ByVal S As Integer, sName As sockaddr, namelen As Integer) As Integer
    Public Declare Function getsockname Lib "Winsock.dll" (ByVal S As Integer, sName As sockaddr, namelen As Integer) As Integer
    Public Declare Function getsockopt Lib "Winsock.dll" (ByVal S As Integer, ByVal Level As Integer, ByVal optname As Integer, optval As Any, optlen As Integer) As Integer
    Public Declare Function htonl Lib "Winsock.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "Winsock.dll" (ByVal hostshort As Integer) As Integer
    Public Declare Function inet_addr Lib "Winsock.dll" (ByVal CP As String) As Long
    Public Declare Function inet_ntoa Lib "Winsock.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "Winsock.dll" (ByVal S As Integer, ByVal backlog As Integer) As Integer
    Public Declare Function ntohl Lib "Winsock.dll" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "Winsock.dll" (ByVal netshort As Integer) As Integer
    Public Declare Function recv Lib "Winsock.dll" (ByVal S As Integer, ByVal buf As Any, ByVal buflen As Integer, ByVal FLAGS As Integer) As Integer
    Public Declare Function recvfrom Lib "Winsock.dll" (ByVal S As Integer, buf As Any, ByVal buflen As Integer, ByVal FLAGS As Integer, from As sockaddr, fromlen As Integer) As Integer
    Public Declare Function ws_select Lib "Winsock.dll" Alias "select" (ByVal nfds As Integer, readfds As Any, writefds As Any, exceptfds As Any, timeout As timeval) As Integer
    Public Declare Function Send Lib "Winsock.dll" Alias "send" (ByVal S As Integer, buf As Any, ByVal buflen As Integer, ByVal FLAGS As Integer) As Integer
    Public Declare Function sendto Lib "Winsock.dll" (ByVal S As Integer, buf As Any, ByVal buflen As Integer, ByVal FLAGS As Integer, to_addr As sockaddr, ByVal tolen As Integer) As Integer
    Public Declare Function setsockopt Lib "Winsock.dll" (ByVal S As Integer, ByVal Level As Integer, ByVal optname As Integer, optval As Any, ByVal optlen As Integer) As Integer
    Public Declare Function ShutDown Lib "Winsock.dll" Alias "shutdown" (ByVal S As Integer, ByVal how As Integer) As Integer
    Public Declare Function Socket Lib "Winsock.dll" Alias "socket" (ByVal af As Integer, ByVal s_type As Integer, ByVal protocol As Integer) As Integer
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "Winsock.dll" (addr As Long, ByVal addr_len As Integer, ByVal addr_type As Integer) As Long
    Public Declare Function gethostbyname Lib "Winsock.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "Winsock.dll" (ByVal host_name As String, ByVal namelen As Integer) As Integer
    Public Declare Function getservbyport Lib "Winsock.dll" (ByVal Port As Integer, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "Winsock.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "Winsock.dll" (ByVal proto As Integer) As Long
    Public Declare Function getprotobyname Lib "Winsock.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "Winsock.dll" (ByVal wVR As Integer, lpWSAD As WSADataType) As Integer
    Public Declare Function WSCleanUp Lib "Winsock.dll" () As Integer
    Public Declare Sub WSASetLastError Lib "Winsock.dll" (ByVal iError As Integer)
    Public Declare Function WSAGetLastError Lib "Winsock.dll" () As Integer
    Public Declare Function WSAIsBlocking Lib "Winsock.dll" () As Integer
    Public Declare Function WSAUnhookBlockingHook Lib "Winsock.dll" () As Integer
    Public Declare Function WSASetBlockingHook Lib "Winsock.dll" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "Winsock.dll" () As Integer
    Public Declare Function WSAAsyncGetServByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetServByPort Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal Port As Integer, ByVal proto As String, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetProtoByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal proto_name As String, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetProtoByNumber Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal number As Integer, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetHostByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal host_name As String, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSAAsyncGetHostByAddr Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, addr As Long, ByVal addr_len As Integer, ByVal addr_type As Integer, buf As Any, ByVal buflen As Integer) As Integer
    Public Declare Function WSACancelAsyncRequest Lib "Winsock.dll" (ByVal hAsyncTaskHandle As Integer) As Integer
    Public Declare Function WSAAsyncSelect Lib "Winsock.dll" (ByVal S As Integer, ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal lEvent As Long) As Integer
    Public Declare Function WSARecvEx Lib "Winsock.dll" (ByVal S As Integer, buf As Any, ByVal buflen As Integer, ByVal FLAGS As Integer) As Integer
#ElseIf Win32 Then
'---Windows System Functions
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
    Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'---async notification constants
    Public Const SOL_SOCKET = &HFFFF&
    Public Const SO_LINGER = &H80&
    Public Const FD_READ = &H1&
    Public Const FD_WRITE = &H2&
    Public Const FD_OOB = &H4&
    Public Const FD_ACCEPT = &H8&
    Public Const FD_CONNECT = &H10&
    Public Const FD_CLOSE = &H20&
'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, addrlen As Long) As Long
    Public Declare Function bind Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function closesocket Lib "wsock32.dll" (ByVal S As Long) As Long
    Public Declare Function Connect Lib "wsock32.dll" Alias "connect" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal S As Long, ByVal cmd As Long, argp As Long) As Long
    Public Declare Function getpeername Lib "wsock32.dll" (ByVal S As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockname Lib "wsock32.dll" (ByVal S As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Public Declare Function inet_addr Lib "wsock32.dll" (ByVal CP As String) As Long
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "wsock32.dll" (ByVal S As Long, ByVal backlog As Long) As Long
    Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Public Declare Function recv Lib "wsock32.dll" (ByVal S As Long, ByVal buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
    Public Declare Function recvfrom Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long, from As sockaddr, fromlen As Long) As Long
    Public Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
    Public Declare Function Send Lib "wsock32.dll" Alias "send" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
    Public Declare Function sendto Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
    Public Declare Function setsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
    Public Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
    Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Public Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Public Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
    Public Declare Function WSCleanUp Lib "wsock32.dll" () As Long
    Public Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)
    Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Public Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Public Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal number As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Public Declare Function WSARecvEx Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
#End If

'Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SW_NOSIZE = 1
Public Const SW_NOMOVE = 2
Public Const FLAGS = SW_NOSIZE Or SW_NOMOVE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public WinsockStartedUp As Boolean     'Tracks wether winsock has been started or not

Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If
End Function

Public Function WSAGetAsyncError(ByVal lParam As Long) As Integer
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

Public Function GetAscIP(ByVal inn As Long) As String
    #If Win32 Then
        Dim nStr&
    #Else
        Dim nStr%
    #End If
    Dim lpStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        MemCopy ByVal retString, ByVal lpStr, nStr
        retString = Left(retString, nStr)
        GetAscIP = retString
    Else
        GetAscIP = "255.255.255.255"
    End If
End Function

#If Win16 Then
    Public Function GetPeerAddress(ByVal S%) As String
    Dim addrlen%
#ElseIf Win32 Then
    Public Function GetPeerAddress(ByVal S&) As String
    Dim addrlen&
#End If
    Dim sa As sockaddr
    addrlen = sockaddr_size
    If getpeername(S, sa, addrlen) Then
        GetPeerAddress = ""
    Else
        GetPeerAddress = SockAddressToString(sa)
    End If
End Function

'this function DOES work on 16 and 32 bit systems
#If Win16 Then
    Function GetSockAddress(ByVal S%) As String
    Dim addrlen%
    Dim ret%
#ElseIf Win32 Then
    Function GetSockAddress(ByVal S&) As String
    Dim addrlen&
    Dim ret&
#End If
    Dim sa As sockaddr
    Dim szRet$
    szRet = String(32, 0)
    addrlen = sockaddr_size
    If getsockname(S, sa, addrlen) Then
        GetSockAddress = ""
    Else
        GetSockAddress = SockAddressToString(sa)
    End If
End Function

'this function should work on 16 and 32 bit systems
Function GetWSAErrorString(ByVal errnum&) As String
    On Error Resume Next
    Select Case errnum
        Case 10004: GetWSAErrorString = "Interrupted system call."
        Case 10009: GetWSAErrorString = "Bad file number."
        Case 10013: GetWSAErrorString = "Permission Denied."
        Case 10014: GetWSAErrorString = "Bad Address."
        Case 10022: GetWSAErrorString = "Invalid Argument."
        Case 10024: GetWSAErrorString = "Too many open files."
        Case 10035: GetWSAErrorString = "Operation would block."
        Case 10036: GetWSAErrorString = "Operation now in progress."
        Case 10037: GetWSAErrorString = "Operation already in progress."
        Case 10038: GetWSAErrorString = "Socket operation on nonsocket."
        Case 10039: GetWSAErrorString = "Destination address required."
        Case 10040: GetWSAErrorString = "Message too long."
        Case 10041: GetWSAErrorString = "Protocol wrong type for socket."
        Case 10042: GetWSAErrorString = "Protocol not available."
        Case 10043: GetWSAErrorString = "Protocol not supported."
        Case 10044: GetWSAErrorString = "Socket type not supported."
        Case 10045: GetWSAErrorString = "Operation not supported on socket."
        Case 10046: GetWSAErrorString = "Protocol family not supported."
        Case 10047: GetWSAErrorString = "Address family not supported by protocol family."
        Case 10048: GetWSAErrorString = "Address already in use."
        Case 10049: GetWSAErrorString = "Can't assign requested address."
        Case 10050: GetWSAErrorString = "Network is down."
        Case 10051: GetWSAErrorString = "Network is unreachable."
        Case 10052: GetWSAErrorString = "Network dropped connection."
        Case 10053: GetWSAErrorString = "Software caused connection abort."
        Case 10054: GetWSAErrorString = "Connection reset by peer."
        Case 10055: GetWSAErrorString = "No buffer space available."
        Case 10056: GetWSAErrorString = "Socket is already connected."
        Case 10057: GetWSAErrorString = "Socket is not connected."
        Case 10058: GetWSAErrorString = "Can't send after socket shutdown."
        Case 10059: GetWSAErrorString = "Too many references: can't splice."
        Case 10060: GetWSAErrorString = "Connection timed out."
        Case 10061: GetWSAErrorString = "Connection refused."
        Case 10062: GetWSAErrorString = "Too many levels of symbolic links."
        Case 10063: GetWSAErrorString = "File name too long."
        Case 10064: GetWSAErrorString = "Host is down."
        Case 10065: GetWSAErrorString = "No route to host."
        Case 10066: GetWSAErrorString = "Directory not empty."
        Case 10067: GetWSAErrorString = "Too many processes."
        Case 10068: GetWSAErrorString = "Too many users."
        Case 10069: GetWSAErrorString = "Disk quota exceeded."
        Case 10070: GetWSAErrorString = "Stale NFS file handle."
        Case 10071: GetWSAErrorString = "Too many levels of remote in path."
        Case 10091: GetWSAErrorString = "Network subsystem is unusable."
        Case 10092: GetWSAErrorString = "Winsock DLL cannot support this application."
        Case 10093: GetWSAErrorString = "Winsock not initialized."
        Case 10101: GetWSAErrorString = "Disconnect."
        Case 11001: GetWSAErrorString = "Host not found."
        Case 11002: GetWSAErrorString = "Nonauthoritative host not found."
        Case 11003: GetWSAErrorString = "Nonrecoverable error."
        Case 11004: GetWSAErrorString = "Valid name, no data record of requested type."
        Case Else:
    End Select
End Function

#If Win16 Then
Public Function SendData(ByVal S%, vMessage As Variant) As Integer
#ElseIf Win32 Then
Public Function SendData(ByVal S&, vMessage As Variant) As Long
#End If
    Dim TheMsg() As Byte, sTemp$
    TheMsg = ""
    Select Case VarType(vMessage)
        Case 8209   'byte array
            sTemp = vMessage
            TheMsg = sTemp
        Case 8      'string, if we recieve a string, its assumed we are linemode
            #If Win32 Then
                sTemp = StrConv(vMessage, vbFromUnicode)
            #Else
                sTemp = vMessage
            #End If
        Case Else
            sTemp = CStr(vMessage)
            #If Win32 Then
                sTemp = StrConv(vMessage, vbFromUnicode)
            #Else
                sTemp = vMessage
            #End If
    End Select
    TheMsg = sTemp
    If UBound(TheMsg) > -1 Then
        SendData = Send(S, TheMsg(0), UBound(TheMsg) + 1, 0)
    End If
End Function

Public Function SockAddressToString(sa As sockaddr) As String
    SockAddressToString = GetAscIP(sa.sin_addr) & ":" & ntohs(sa.sin_port)
End Function

'returns IP as long, in network byte order
Public Function GetHostByNameAlias(ByVal hostname$) As Long
    'Return IP address as a long, in network byte order
    Dim phe&
    Dim heDestHost As HostEnt
    Dim addrList&
    Dim retIP&
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
End Function
