Attribute VB_Name = "WinInet"
Option Explicit

' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Public Const scUserAgent = "http sample"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000

' Security constants
Public Const INTERNET_OPTION_SECURITY_FLAGS = 31
Public Const SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
Public Const INTERNET_FLAG_SECURE = &H800000


' Sends the specified request to the HTTP server.
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_STATUS_DENIED = 401
Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000
' Reads data from a handle opened by the HttpOpenRequest function.
Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

'------------------------------------------------------------------------------------
'函数:httpSend  发送http请求包,不能用于发送https请求包
'参数:（Interner会话句柄,目标服务器,目标文件,[模式],[Referer],[头信息],[data数据体]）
'返回值:返回一个在HttpQueryInfo使用的句柄
'------------------------------------------------------------------------------------
Function httpSend(ByVal hInternetSession As Long, ByVal szHost As String, ByVal szObject As String, _
                  Optional ByVal szVerb As String = "GET", Optional ByVal szReferer As String = vbNullString, _
                  Optional ByVal szHeaders As String = vbNullString, Optional ByVal szData As String = vbNullString, _
                  Optional ByVal ishttps As Boolean = False) As Long
    Dim szAddHeader As String
    Dim dwPort As Long
    Dim SecFlag As Long
    Dim dwTimeOut As Long
    Dim headerLength As Long
    Dim hInternetConnect As Long
    '----------------------------支持https-------------------------
    If ishttps = True Then
        dwPort = INTERNET_DEFAULT_HTTPS_PORT
        SecFlag = INTERNET_FLAG_SECURE Or INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
    Else
        dwPort = INTERNET_DEFAULT_HTTP_PORT
        SecFlag = 0
    End If
    '----------------------------获取句柄-------------------------
    hInternetConnect = InternetConnect(hInternetSession, szHost, dwPort, "codiy", vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    httpSend = HttpOpenRequest(hInternetConnect, szVerb, szObject, "HTTP/1.0", szReferer, 0, INTERNET_FLAG_NO_AUTO_REDIRECT Or INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_RELOAD Or SecFlag, 0)
    '---------------------------添加头部--------------------------
    szAddHeader = "Accept-Language: zh-CN,zh;q=0.8" & vbCrLf
    HttpAddRequestHeaders httpSend, szAddHeader, Len(szAddHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD
    szAddHeader = "Connection: Keep-Alive" & vbCrLf
    HttpAddRequestHeaders httpSend, szAddHeader, Len(szAddHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD
    If szHeaders <> vbNullString Then
        Do While Len(szHeaders) <> 1
            headerLength = InStr(szHeaders, vbCrLf)
            szAddHeader = Left(szHeaders, headerLength + 1)
            szHeaders = Mid(szHeaders, headerLength + 1)
            HttpAddRequestHeaders httpSend, szAddHeader, Len(szAddHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD
        Loop
    Else
        szAddHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
        HttpAddRequestHeaders httpSend, szAddHeader, Len(szAddHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD
    End If
    '---------------------------设置超时--------------------------
    dwTimeOut = 180000 ' time out is set to 3 minutes
    InternetSetOption httpSend, INTERNET_OPTION_CONNECT_TIMEOUT, dwTimeOut, 4
    InternetSetOption httpSend, INTERNET_OPTION_RECEIVE_TIMEOUT, dwTimeOut, 4
    InternetSetOption httpSend, INTERNET_OPTION_SEND_TIMEOUT, dwTimeOut, 4
    '---------------------------发送数据--------------------------
    If HttpSendRequest(httpSend, vbNullString, 0, szData, Len(szData)) = 0 Then
    MsgBox "HttpSendRequest fault"
    End If
End Function

'------------------------------------------------------------------------------------
'函数:httpReadData 读取http响应包中的data数据（用HttpQueryInfo读取http响应包中的头信息）
'参数:（句柄）由HttpOpenRequest或者httpSend函数返回的句柄
'返回值:获取的data数据
'------------------------------------------------------------------------------------
Function httpReadData(ByVal hInternet As String) As String
Dim bDoLoop             As Boolean
Dim sReadBuffer         As String * 2048
Dim lNumberOfBytesRead  As Long
bDoLoop = True
While bDoLoop
    sReadBuffer = vbNullString
    bDoLoop = InternetReadFile(hInternet, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
    httpReadData = httpReadData & Left$(sReadBuffer, lNumberOfBytesRead)
    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
Wend
End Function
