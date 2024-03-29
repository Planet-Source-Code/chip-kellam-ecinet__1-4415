Attribute VB_Name = "Winsock"
'Visual Basic 4.0 and above winsock declares and functions modules
'requires a msghook for async methods and functions. This declares
'file was originally obtained from the alt.winsock.programming news
'group, alot has been added and modified since that time. However I
'would like to credit the people on that newsgroup for the information
'they contributed, and now I pass it back...
'
'NOTES:
' I haven't been able to get the WSAAsyncGetXbyY functions to work properly
'under windows95(tm) aside from that ALL functions "SHOULD" work just fine.
'any questions about this file may be posted to alt.winsock.programming
'   Topaz..
'
Option Explicit

'windows declares here
#If Win16 Then
    Declare Function PostMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
    Declare Sub MemCopy Lib "Kernel" Alias "hmemcpy" (Dest As Any, Src As Any, ByVal cb&)
    Declare Function lstrlen Lib "Kernel" (ByVal lpString As Any) As Integer
#ElseIf Win32 Then
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
    Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
#End If
'WINSOCK DEFINES START HERE
'not same
Global Const FD_SETSIZE = 64
Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type

'same
Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Const hostent_size = 16

Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type
Const servent_size = 14

Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type
Const protoent_size = 10

Global Const IPPROTO_TCP = 6
Global Const IPPROTO_UDP = 17

Global Const INADDR_NONE = &HFFFF
Global Const INADDR_ANY = &H0

Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Const sockaddr_size = 16
Dim saZero As sockaddr


Global Const WSA_DESCRIPTIONLEN = 256
Global Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Global Const WSA_SYS_STATUS_LEN = 128
Global Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Global Const INVALID_SOCKET = -1
Global Const SOCKET_ERROR = -1

Global Const SOCK_STREAM = 1
Global Const SOCK_DGRAM = 2

Global Const MAXGETHOSTSTRUCT = 1024

Global Const AF_INET = 2
Global Const PF_INET = 2

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

#If Win32 Then
    Global Const SOL_SOCKET = &HFFFF&
    Global Const SO_LINGER = &H80&
    Global Const FD_READ = &H1&
    Global Const FD_WRITE = &H2&
    Global Const FD_OOB = &H4&
    Global Const FD_ACCEPT = &H8&
    Global Const FD_CONNECT = &H10&
    Global Const FD_CLOSE = &H20&
    Global Const MSG_PEEK = &H2&
#Else
    Global Const SOL_SOCKET = &HFFFF
    Global Const SO_LINGER = &H80
    Global Const FD_READ = &H1
    Global Const FD_WRITE = &H2
    Global Const FD_OOB = &H4
    Global Const FD_ACCEPT = &H8
    Global Const FD_CONNECT = &H10
    Global Const FD_CLOSE = &H20
    Global Const MSG_PEEK = &H2
#End If



'SOCKET FUNCTIONS
#If Win16 Then
'SOCKET FUNCTIONS
    Declare Function accept Lib "Winsock.dll" (ByVal s As Integer, addr As sockaddr, addrlen As Integer) As Integer
    Declare Function bind Lib "Winsock.dll" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
    Declare Function closesocket Lib "Winsock.dll" (ByVal s As Integer) As Integer
    Declare Function connect Lib "Winsock.dll" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
    Declare Function ioctlsocket Lib "Winsock.dll" (ByVal s As Integer, ByVal cmd As Long, argp As Long) As Integer
    Declare Function getpeername Lib "Winsock.dll" (ByVal s As Integer, sname As sockaddr, namelen As Integer) As Integer
    Declare Function getsockname Lib "Winsock.dll" (ByVal s As Integer, sname As sockaddr, namelen As Integer) As Integer
    Declare Function getsockopt Lib "Winsock.dll" (ByVal s As Integer, ByVal level As Integer, ByVal optname As Integer, ByVal optval As String, optlen As Integer) As Integer
    Declare Function htonl Lib "Winsock.dll" (ByVal hostlong As Long) As Long
    Declare Function htons Lib "Winsock.dll" (ByVal hostshort As Integer) As Integer
    Declare Function inet_addr Lib "Winsock.dll" (ByVal cp As String) As Long
    Declare Function inet_ntoa Lib "Winsock.dll" (ByVal inn As Long) As Long
    Declare Function listen Lib "Winsock.dll" (ByVal s As Integer, ByVal backlog As Integer) As Integer
    Declare Function ntohl Lib "Winsock.dll" (ByVal netlong As Long) As Long
    Declare Function ntohs Lib "Winsock.dll" (ByVal netshort As Integer) As Integer
    Declare Function recv Lib "Winsock.dll" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
    Declare Function recvfrom Lib "Winsock.dll" (ByVal s As Integer, ByVal buf As String, ByVal buflen As Integer, ByVal flags As Integer, from As sockaddr, fromlen As Integer) As Integer
    Declare Function ws_select Lib "Winsock.dll" Alias "select" (ByVal nfds As Integer, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Integer
    Declare Function Send Lib "Winsock.dll" Alias "send" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
    Declare Function sendto Lib "Winsock.dll" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer, to_addr As sockaddr, ByVal tolen As Integer) As Integer
    Declare Function setsockopt Lib "Winsock.dll" (ByVal s As Integer, ByVal level As Integer, ByVal optname As Integer, optval As Any, ByVal optlen As Integer) As Integer
    Declare Function ShutDown Lib "Winsock.dll" Alias "shutdown" (ByVal s As Integer, ByVal how As Integer) As Integer
    Declare Function socket Lib "Winsock.dll" (ByVal af As Integer, ByVal s_type As Integer, ByVal protocol As Integer) As Integer
'DATABASE FUNCTIONS
    Declare Function gethostbyaddr Lib "Winsock.dll" (addr As Long, ByVal addr_len As Integer, ByVal addr_type As Integer) As Long
    Declare Function gethostbyname Lib "Winsock.dll" (ByVal host_name As String) As Long
    Declare Function gethostname Lib "Winsock.dll" (ByVal host_name As String, ByVal namelen As Integer) As Integer
    Declare Function getservbyport Lib "Winsock.dll" (ByVal Port As Integer, ByVal proto As String) As Long
    Declare Function getservbyname Lib "Winsock.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Declare Function getprotobynumber Lib "Winsock.dll" (ByVal proto As Integer) As Long
    Declare Function getprotobyname Lib "Winsock.dll" (ByVal proto_name As String) As Long
'WINDOWS EXTENSIONS
    Declare Function WSAStartup Lib "Winsock.dll" (ByVal wVR As Integer, lpWSAD As WSADataType) As Integer
    Declare Function WSACleanup Lib "Winsock.dll" () As Integer
    Declare Sub WSASetLastError Lib "Winsock.dll" (ByVal iError As Integer)
    Declare Function WSAGetLastError Lib "Winsock.dll" () As Integer
    Declare Function WSAIsBlocking Lib "Winsock.dll" () As Integer
    Declare Function WSAUnhookBlockingHook Lib "Winsock.dll" () As Integer
    Declare Function WSASetBlockingHook Lib "Winsock.dll" (ByVal lpBlockFunc As Long) As Long
    Declare Function WSACancelBlockingCall Lib "Winsock.dll" () As Integer
    Declare Function WSAAsyncGetServByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal serv_name As String, ByVal proto As String, ByVal buf As String, ByVal buflen As Integer) As Integer
    Declare Function WSAAsyncGetServByPort Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal Port As Integer, ByVal proto As String, ByVal buf As String, ByVal buflen As Integer) As Integer
    Declare Function WSAAsyncGetProtoByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal proto_name As String, ByVal buf As String, ByVal buflen As Integer) As Integer
    Declare Function WSAAsyncGetProtoByNumber Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal number As Integer, ByVal buf As String, ByVal buflen As Integer) As Integer
    Declare Function WSAAsyncGetHostByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal host_name As String, ByVal buf As String, ByVal buflen As Integer) As Integer
    Declare Function WSAAsyncGetHostByAddr Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal addr As String, ByVal addr_len As Integer, ByVal addr_type As Integer, ByVal buf As String, ByVal buflen As Integer) As Integer
    Declare Function WSACancelAsyncRequest Lib "Winsock.dll" (ByVal hAsyncTaskHandle As Integer) As Integer
    Declare Function WSAAsyncSelect Lib "Winsock.dll" (ByVal s As Integer, ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal lEvent As Long) As Integer
    Declare Function WSARecvEx Lib "Winsock.dll" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
#ElseIf Win32 Then
'SOCKET FUNCTIONS
    Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
    Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
    Declare Function connect Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Declare Function ioctlsocket Lib "wsock32.dll" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
    Declare Function getpeername Lib "wsock32.dll" (ByVal s As Long, sname As sockaddr, namelen As Long) As Long
    Declare Function getsockname Lib "wsock32.dll" (ByVal s As Long, sname As sockaddr, namelen As Long) As Long
    Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByVal optval As String, optlen As Long) As Long
    Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
    Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Declare Function recv Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Declare Function recvfrom Lib "wsock32.dll" (ByVal s As Long, ByVal buf As String, ByVal buflen As Long, ByVal flags As Long, from As sockaddr, fromlen As Long) As Long
    Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
    Declare Function Send Lib "wsock32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Declare Function sendto Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
    Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
    Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
    Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'DATABASE FUNCTIONS
    Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
'WINDOWS EXTENSIONS
    Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
    Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)
    Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
'WSAASYNCGETXBYY FUNCTIONS DON'T WORK RELIABLY UNDER 32BIT WINDOWS
    Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, ByVal buf As String, ByVal buflen As Long) As Long
    Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, ByVal buf As String, ByVal buflen As Long) As Long
    Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, ByVal buf As String, ByVal buflen As Long) As Long
    Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal number As Long, ByVal buf As String, ByVal buflen As Long) As Long
    Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, ByVal buf As String, ByVal buflen As Long) As Long
    Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal addr As String, ByVal addr_len As Long, ByVal addr_type As Long, ByVal buf As String, ByVal buflen As Long) As Long
    Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Declare Function WSARecvEx Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
#End If




'SOME STUFF I ADDED
Global Const INVALID_PORT = -1  'added by me
Global Const INVALID_PROTO = -1 'added by me

Global MySocket%
Global SockReadBuffer$

Global Const WSA_NoName = "Unknown"

Global WSAStartedUp%    'Flag to keep track of whether winsock WSAStartup wascalled

Global WSARecvActive%   'Flag to indicate an async recv is in progress
Global WSASendActive%   'Flag to indicate an async send is in progress

'these are old functions, or examples of various things
'
'
Private Sub oldfuncs()
'use this function for the read function in line mode apps
'    Dim s$, buf$, a%, i%, ii%
'    buf$ = String$(1024, " ")
'    a% = recv%(MySocket%, ByVal buf$, 1024, 0)
'    If a% > 0 Then
'        'SockReadBuffer$ = SockReadBuffer$ + RTrim$(buf$)
'        SockReadBuffer$ = SockReadBuffer$ + Left$(buf$, a%)
'        While InStr(SockReadBuffer$, EndLine$)
'            i = InStr(SockReadBuffer$, EndLine$)
'            If i Then
'                If i < Len(SockReadBuffer$) Then
'                    s$ = Left$(SockReadBuffer$, i - 1)
'                    SockReadBuffer$ = Mid$(SockReadBuffer$, i + 1)
'                    If InStr(s$, Chr$(13)) Then
'                        s$ = Left$(s$, InStr(s$, Chr$(13)) - 1)
'                    ElseIf InStr(s$, Chr$(10)) Then
'                        s$ = Left$(s$, InStr(s$, Chr$(10)) - 1)
'                    End If
'                    ServerHandler s$ & CrLf$
'                    'Debug.Print "|"; s$; "|"
'                Else
'                    s$ = SockReadBuffer$
'                    SockReadBuffer$ = ""
'                    If InStr(s$, Chr$(13)) Then
'                        s$ = Left$(s$, InStr(s$, Chr$(13)) - 1)
'                    ElseIf InStr(s$, Chr$(10)) Then
'                        s$ = Left$(s$, InStr(s$, Chr$(10)) - 1)
'                    End If
'                    ServerHandler s$ & CrLf$
'                    'Debug.Print "|"; s$; "|"
'                End If
'            End If
'        Wend
'    End If

End Sub

'this function should work on 16 and 32 bit systems
Function WSAGetAsyncBufLen(ByVal lParam As Long) As Long
'used only by the async lookups and wparam will
'be the task handle for these events
    On Error Resume Next
    '#define WSAGETASYNCBUFLEN(lParam)           LOWORD(lParam)
    WSAGetAsyncBufLen = lParam And &HFFFF&
    If Err Then
        WSAGetAsyncBufLen = 0
    End If
End Function

'this function should work on 16 and 32 bit systems
Function WSAGetSelectEvent(ByVal lParam As Long) As Long
    On Error Resume Next
    '#define WSAGETSELECTEVENT(lParam)            LOWORD(lParam)
    WSAGetSelectEvent = (lParam And &HFFFF&)
    If Err Then
        WSAGetSelectEvent = 0
    End If
End Function



'this function should work on 16 and 32 bit systems
Function WSAGetAsyncError(ByVal lParam As Long) As Long
    On Error Resume Next
    'WSAGETASYNCERROR(lParam) HIWORD(lParam)
    WSAGetAsyncError = (lParam \ &H10000 And &HFFFF&)
    If Err Then
        WSAGetAsyncError = 0
    End If
End Function



'this function DOES work on 16 and 32 bit systems
Function AddrToIP(ByVal AddrOrIP$) As String
    On Error Resume Next
    AddrToIP$ = getascip(GetHostByNameAlias(AddrOrIP$))
    If Err Then AddrToIP$ = "255.255.255.255"
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function ConnectSock(ByVal host$, ByVal Port$, ByVal HWndToMsg%, ByVal Async%) As Integer
    Dim s%, SelectOps%, dummy%
#ElseIf Win32 Then
    Function ConnectSock(ByVal host$, ByVal Port$, ByVal HWndToMsg&, ByVal Async%) As Long
    Dim s&, SelectOps&, dummy&
#End If
    Dim sockin As sockaddr
   
    SockReadBuffer$ = ""
    sockin = saZero
    sockin.sin_family = AF_INET

    sockin.sin_port = GetServiceByName(Port$, "tcp")
    If sockin.sin_port = INVALID_PORT Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    
    sockin.sin_addr = GetHostByNameAlias(host$)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
'    retIpPort$ = getascip$(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)

    s = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If SetSockLinger(s, 1, 0) = SOCKET_ERROR Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If Not Async Then
        If connect(s, sockin, sockaddr_size) <> 0 Then
            If s > 0 Then
                dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
'        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
'        If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
'            If s > 0 Then
'                dummy = closesocket(s)
'            End If
'            ConnectSock = INVALID_SOCKET
'            Exit Function
'        End If
    Else
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
            If s > 0 Then
                dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        If connect(s, sockin, sockaddr_size) <> -1 Then
            If s > 0 Then
                dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    End If
    ConnectSock = s
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function SetSockLinger(ByVal SockNum%, ByVal OnOff%, ByVal LingerTime%) As Integer
    Dim ret%, ret1%
#ElseIf Win32 Then
    Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
    Dim ret&, ret1&
#End If
    Dim LingBuf$, Linger As LingerType
    LingBuf = Space(4)
    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    MemCopy ByVal LingBuf, Linger, 4
    ret = setsockopt(SockNum, SOL_SOCKET, SO_LINGER, ByVal LingBuf, 4)
    Debug.Print "Error Setting Linger info: "; ret
    
    LingBuf = Space(4)
    Linger.l_onoff = 0   'reset this so we can get an honest look
    Linger.l_linger = 0  'reset this so we can get an honest look
    
    ret = getsockopt(SockNum, SOL_SOCKET, SO_LINGER, LingBuf, 4)
    MemCopy Linger, ByVal LingBuf, 4
    Debug.Print "Error Getting Linger info: "; ret
    Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
    Debug.Print "Linger time if linger is on: "; Linger.l_linger
    SetSockLinger = ret
End Function

'this function DOES work on 16 and 32 bit systems
Sub EndWinsock()
    Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub

'this function DOES work on 16 and 32 bit systems
Function getascip(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr&
#If Win16 Then
    Dim nStr%
#ElseIf Win32 Then
    Dim nStr&
#End If
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function

'this function DOES work on 32bit and 16 bit systems
Function GetHostByAddress(ByVal addr As Long) As String
    On Error Resume Next
    Dim phe&, ret&
    Dim heDestHost As HostEnt
    Dim hostname$
    phe = gethostbyaddr(addr, 4, PF_INET)
    Debug.Print phe
    If phe <> 0 Then
        MemCopy heDestHost, ByVal phe, hostent_size
        Debug.Print heDestHost.h_name
        Debug.Print heDestHost.h_aliases
        Debug.Print heDestHost.h_addrtype
        Debug.Print heDestHost.h_length
        Debug.Print heDestHost.h_addr_list

        hostname = String(256, 0)
        MemCopy ByVal hostname, ByVal heDestHost.h_name, 256
        GetHostByAddress = Left(hostname, InStr(hostname, Chr(0)) - 1)
    Else
        GetHostByAddress = WSA_NoName
    End If
    If Err Then GetHostByAddress = WSA_NoName
End Function

'this function DOES work on 16 and 32 bit systems
Function GetHostByNameAlias(ByVal hostname$) As Long
    On Error Resume Next
    'Return IP address as a long, in network byte order

    Dim phe&    ' pointer to host information entry
    Dim heDestHost As HostEnt 'hostent structure
    Dim addrList&
    Dim retIP&
    'first check to see if what we have been passed is a valid IP
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        'it wasn't an IP, so do a DNS lookup
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            'Pointer is non-null, so copy in hostent structure
            MemCopy heDestHost, ByVal phe, hostent_size
            'Now get first pointer in address list
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            'its not a valid address
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function

'this function DOES work on 16 and 32 bit systems
Function GetLocalHostName() As String
    Dim dummy&
    Dim LocalName$
    Dim s$
    On Error Resume Next
    LocalName = String(256, 0)
    LocalName = WSA_NoName
    dummy = 1
    s = String(256, 0)
    dummy = gethostname(s, 256)
    If dummy = 0 Then
        s = Left(s, InStr(s, Chr(0)) - 1)
        If Len(s) > 0 Then
            LocalName = s
        End If
    End If
    GetLocalHostName = LocalName
    If Err Then GetLocalHostName = WSA_NoName
End Function

'this function DOES work on 16 and 32 bit systems
#If Win16 Then
    Function GetPeerAddress(ByVal s%) As String
    Dim addrlen%
    Dim ret%
#ElseIf Win32 Then
    Function GetPeerAddress(ByVal s&) As String
    Dim addrlen&
    Dim ret&
#End If
    On Error Resume Next
    Dim sa As sockaddr
    addrlen = sockaddr_size
    ret = getpeername(s, sa, addrlen)
    If ret = 0 Then
        GetPeerAddress = SockaddressToString(sa)
    Else
        GetPeerAddress = ""
    End If
    If Err Then GetPeerAddress = ""
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function GetPortFromString(ByVal PortStr$) As Integer
#ElseIf Win32 Then
    Function GetPortFromString(ByVal PortStr$) As Long
#End If
    'sometimes users provide ports outside the range of a VB
    'integer, so this function returns an integer for a string
    'just to keep an error from happening, it converts the
    'number to a negative if needed
    On Error Resume Next
    If Val(PortStr$) > 32767 Then
        GetPortFromString = CInt(Val(PortStr$) - &H10000)
    Else
        GetPortFromString = Val(PortStr$)
    End If
    If Err Then GetPortFromString = 0
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function GetProtocolByName(ByVal protocol$) As Integer
    Dim tmpShort%
#ElseIf Win32 Then
    Function GetProtocolByName(ByVal protocol$) As Long
    Dim tmpShort&
#End If
    On Error Resume Next
    Dim ppe&
    Dim peDestProt As protoent
    ppe = getprotobyname(protocol)
    If ppe = 0 Then
        tmpShort = Val(protocol)
        If tmpShort <> 0 Or protocol = "0" Or protocol = "" Then
            GetProtocolByName = htons(tmpShort)
        Else
            GetProtocolByName = INVALID_PROTO
        End If
    Else
        MemCopy peDestProt, ByVal ppe, protoent_size
        GetProtocolByName = peDestProt.p_proto
    End If
    If Err Then GetProtocolByName = INVALID_PROTO
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function GetServiceByName(ByVal service$, ByVal protocol$) As Integer
    Dim serv%
#ElseIf Win32 Then
    Function GetServiceByName(ByVal service$, ByVal protocol$) As Long
    Dim serv&
#End If
    On Error Resume Next
    Dim pse&
    Dim seDestServ As servent
    pse = getservbyname(service, protocol)
    If pse <> 0 Then
        MemCopy seDestServ, ByVal pse, servent_size
        GetServiceByName = seDestServ.s_port
    Else
        serv = Val(service)
        If serv <> 0 Then
            GetServiceByName = htons(serv)
        Else
            GetServiceByName = INVALID_PORT
        End If
    End If
    If Err Then GetServiceByName = INVALID_PORT
End Function

'this function DOES work on 16 and 32 bit systems
#If Win16 Then
    Function GetSockAddress(ByVal s%) As String
    Dim addrlen%
    Dim ret%
#ElseIf Win32 Then
    Function GetSockAddress(ByVal s&) As String
    Dim addrlen&
    Dim ret&
#End If
    On Error Resume Next
    Dim sa As sockaddr
    Dim szRet$
    szRet = String(32, 0)
    addrlen = sockaddr_size
    ret = getsockname(s, sa, addrlen)
    If ret = 0 Then
        GetSockAddress = SockaddressToString(sa)
    Else
        GetSockAddress = ""
    End If
    If Err Then GetSockAddress = ""
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

'this function DOES work on 16 and 32 bit systems
Function IpToAddr(ByVal AddrOrIP$) As String
    On Error Resume Next
    IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP$))
    If Err Then IpToAddr = WSA_NoName
End Function

'this function DOES work on 16 and 32 bit systems
Function IrcGetAscIp(ByVal IPL$) As String
    'this function is IRC specific, it expects a long ip stored in Network byte order, in a string
    'the kind that would be parsed out of a DCC command string
    On Error GoTo IrcGetAscIPError:
    Dim lpStr&
#If Win16 Then
    Dim nStr%
#ElseIf Win32 Then
    Dim nStr&
#End If
    Dim retString$
    Dim inn&
    If Val(IPL) > 2147483647 Then
        inn = Val(IPL) - 4294967296#
    Else
        inn = Val(IPL)
    End If
    inn = ntohl(inn)
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        IrcGetAscIp = "0.0.0.0"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    IrcGetAscIp = retString
    Exit Function
IrcGetAscIPError:
    IrcGetAscIp = "0.0.0.0"
    Exit Function
    Resume
End Function

'this function DOES work on 16 and 32 bit systems
Function IrcGetLongIp(ByVal AscIp$) As String
    'this function converts an ascii ip string into a long ip in network byte order
    'and stick it in a string suitable for use in a DCC command.
    On Error GoTo IrcGetLongIpError:
    Dim inn&
    inn = inet_addr(AscIp)
    inn = htonl(inn)
    If inn < 0 Then
        IrcGetLongIp = CVar(inn + 4294967296#)
        Exit Function
    Else
        IrcGetLongIp = CVar(inn)
        Exit Function
    End If
    Exit Function
IrcGetLongIpError:
    IrcGetLongIp = "0"
    Exit Function
    Resume
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function ListenForConnect(ByVal Port%, ByVal HWndToMsg%) As Integer
    Dim s%, dummy%
    Dim SelectOps%
#ElseIf Win32 Then
    Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&) As Long
    Dim s&, dummy&
    Dim SelectOps&
#End If
    Dim sockin As sockaddr
    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_PORT Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    s = socket(PF_INET, SOCK_STREAM, 0)
    If s < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    If bind(s, sockin, sockaddr_size) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = SOCKET_ERROR
        Exit Function
    End If
    
    If listen(s, 1) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    ListenForConnect = s
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function SendData(ByVal s%, ByVal Message$) As Integer
#ElseIf Win32 Then
    Function SendData(ByVal s&, ByVal Message$) As Long
#End If
    SendData = Send(s, ByVal Message, Len(Message), 0)
End Function
'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function RecvLine(ByVal s%) As String
#ElseIf Win32 Then
    Function RecvLine(ByVal s&) As String
#End If
    Dim buf$, l As Integer, eol As Integer
    buf$ = String$(1024, " ")
    l = recv(s, ByVal buf$, 1024, MSG_PEEK)
    If l = 0 Then RecvLine = "": Exit Function
    If l > 0 Then
       eol = InStr(buf$, Chr(10))
       If eol < 1 Then eol = l
       
       l = recv(s, ByVal buf$, eol, 0)
       buf$ = Left$(buf$, l)
'       buf = Replace(sbuf, Chr(13), "")
    Else
       buf$ = ""
    End If
    
    RecvLine = buf$
End Function
#If Win16 Then
    Function WaitForStatus(ByVal s%) As Integer
#ElseIf Win32 Then
    Function WaitForStatus(ByVal s&) As Long
#End If
 
Dim Line$, i, l

   Line$ = RecvLine(s)
   Line$ = Trim(Line$)
   i = 1: l = Len(Line$)
   Do While i <= l
      If Mid(Line$, i, 1) < "0" Or Mid(Line$, i, 1) > "9" Then Exit Do
      i = i + 1
   Loop
If i > 0 Then WaitForStatus = Val(Mid(Line$, 1, i - 1)) Else WaitForStatus = -1
  
End Function


'this function should work on 16 and 32 bit systems
Function SockaddressToString(sa As sockaddr) As String
    On Error Resume Next
    SockaddressToString = getascip(sa.sin_addr) & ":" & ntohs(sa.sin_port)
    If Err Then SockaddressToString = ""
End Function

'this function DOES work on 16 and 32 bit systems
#If Win16 Then
    Function StartWinsock(desc$) As Integer
    Dim ret%
    Dim WinsockVers%
#ElseIf Win32 Then
    Function StartWinsock(desc$) As Long
    Dim ret&
    Dim WinsockVers&
#End If
    
    Dim wsadStartupData As WSADataType
    WinsockVers = &H101   'Vers 1.1

    If WSAStartedUp = False Then

        ret = 1
        ret = WSAStartup(WinsockVers, wsadStartupData)
        If ret = 0 Then
            WSAStartedUp = True
            Debug.Print "wVersion="; VBntoaVers(wsadStartupData.wVersion), "wHighVersion="; VBntoaVers(wsadStartupData.wHighVersion)
            Debug.Print "szDescription="; wsadStartupData.szDescription
            Debug.Print "szSystemStatus="; wsadStartupData.szSystemStatus
            Debug.Print "iMaxSockets="; wsadStartupData.iMaxSockets, "iMaxUdpDg="; wsadStartupData.iMaxUdpDg
            desc = wsadStartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function VBntoaVers$(ByVal vers%)
#ElseIf Win32 Then
    Function VBntoaVers$(ByVal vers&)
#End If
    On Error Resume Next
    Dim szVers$
    szVers = String(5, 0)
    szVers = (vers And &HFF) & "." & ((vers And &HFF00) / 256)
    VBntoaVers = szVers
End Function

'this function should work on 16 and 32 bit systems
'this function uses MemCopy to transfer data from
'2 integers into 2 strings, then combines the strings
'and copys that data into a long, ineffect MAKELONG()
    Function WSAMakeSelectReply(ByVal TheEvent%, ByVal TheError%) As Long
    Dim EventStr$, ErrorStr$, BothStr$, TheLong&
    EventStr = Space(2)
    ErrorStr = Space(2)
    BothStr = Space(4)
    MemCopy ByVal EventStr, TheEvent, 2
    MemCopy ByVal ErrorStr, TheError, 2
    BothStr = EventStr & ErrorStr
    If Len(BothStr) = 4 Then
        MemCopy TheLong, ByVal BothStr, 4
    End If
    WSAMakeSelectReply = TheLong
End Function



Sub Main()

End Sub


