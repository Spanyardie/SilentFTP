Attribute VB_Name = "modWinInet"
Option Explicit

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'seb addition - ftp put context
'we want the filename, gameserverip gameserverport
Type FTP_FILE_PUT_CONTEXT
    sType As String * 60
    sFilename As String * 255
    sServerIP As String * 17
    sServerPort As Long
End Type

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Public Const scUserAgent = "vb wininet"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const INTERNET_FLAG_PASSIVE = &H8000000

'seb addition - active flag
Public Const INTERNET_FLAG_ACTIVE = 0

'valid status callback
Public Const INTERNET_INVALID_STATUS_CALLBACK = -1

'seb addition - where are the callbacks?
Public Declare Function InternetSetStatusCallback Lib "wininet.dll" Alias "InternetSetStatusCallbackA" _
(ByVal hInternetSession As Long, ByVal lpCallbackFunc As Long) As Long

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

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

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
'Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Public Const INTERNET_FLAG_MULTIPART = &H200000

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' // status manifests for Internet status callback
Public Enum InternetStatusFlags
  INTERNET_STATUS_RESOLVING_NAME = 10
  INTERNET_STATUS_NAME_RESOLVED = 11
  INTERNET_STATUS_CONNECTING_TO_SERVER = 20
  INTERNET_STATUS_CONNECTED_TO_SERVER = 21
  INTERNET_STATUS_SENDING_REQUEST = 30
  INTERNET_STATUS_REQUEST_SENT = 31
  INTERNET_STATUS_RECEIVING_RESPONSE = 40
  INTERNET_STATUS_RESPONSE_RECEIVED = 41
  INTERNET_STATUS_CTL_RESPONSE_RECEIVED = 42
  INTERNET_STATUS_PREFETCH = 43
  INTERNET_STATUS_CLOSING_CONNECTION = 50
  INTERNET_STATUS_CONNECTION_CLOSED = 51
  INTERNET_STATUS_HANDLE_CREATED = 60
  INTERNET_STATUS_HANDLE_CLOSING = 70
  INTERNET_STATUS_DETECTING_PROXY = 80
  INTERNET_STATUS_REQUEST_COMPLETE = 100
  INTERNET_STATUS_REDIRECT = 110
  INTERNET_STATUS_INTERMEDIATE_RESPONSE = 120
  INTERNET_STATUS_USER_INPUT_REQUIRED = 140
  INTERNET_STATUS_STATE_CHANGE = 200
End Enum

' // the following can be indicated in a state change notification:
Public Enum InternetStates
  INTERNET_STATE_CONNECTED = &H1&                     ' // connected state (mutually exclusive with disconnected)
  INTERNET_STATE_DISCONNECTED = &H2&                  ' // disconnected from network
  INTERNET_STATE_DISCONNECTED_BY_USER = &H10&         ' // disconnected by user request
  INTERNET_STATE_IDLE = &H100&                        ' // no network requests being made (by Wininet)
  INTERNET_STATE_BUSY = &H200&                        ' // network requests being made (by Wininet)
End Enum

Public Type INTERNET_ASYNC_RESULT
  dwResult          As Long               ' API return value
  dwError           As Long               ' The error code if dwResult is a failure code
End Type

' custom version of INTERNET_ASYNC_RESULT:
Public Type INTERNET_ASYNC_RESULT_X
  Async             As INTERNET_ASYNC_RESULT
  MagicNumber       As Long               ' Always OverlappedMagicNumber
  dwBytesTransfered As Long               ' Counter reported by completion routing
  dwBytesToTransfer As Long               ' Counter reported during Write operations
  AsyncIoPending    As Boolean            ' So caller knows async I/O still pending
  hEvent            As Long
End Type

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
() '    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
() '(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
() '(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
() '    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
' Initializes an application's use of the Win32 Internet functions

' Opens an HTTP request handle.
'Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
'(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
'ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Sends the specified request to the HTTP server.
'Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
'hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
'String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
'Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
'(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
'ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' The possible values for the lInfoLevel parameter include:
'Public Const HTTP_QUERY_CONTENT_TYPE = 1
'Public Const HTTP_QUERY_CONTENT_LENGTH = 5
'Public Const HTTP_QUERY_EXPIRES = 10
'Public Const HTTP_QUERY_LAST_MODIFIED = 11
'Public Const HTTP_QUERY_PRAGMA = 17
'Public Const HTTP_QUERY_VERSION = 18
'Public Const HTTP_QUERY_STATUS_CODE = 19
'Public Const HTTP_QUERY_STATUS_TEXT = 20
'Public Const HTTP_QUERY_RAW_HEADERS = 21
'Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
'Public Const HTTP_QUERY_FORWARDED = 30
'Public Const HTTP_QUERY_SERVER = 37
'Public Const HTTP_QUERY_USER_AGENT = 39
'Public Const HTTP_QUERY_SET_COOKIE = 43
'Public Const HTTP_QUERY_REQUEST_METHOD = 45
'Public Const HTTP_STATUS_DENIED = 401
'Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407
'
' Add this flag to the about flags to get request header.
'Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
'Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
'Public Declare Function InternetReadFile Lib "wininet.dll" _
'(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
'lNumberOfBytesRead As Long) As Integer
'
'Public Declare Function InternetWriteFile Lib "wininet.dll" _
'        (ByVal hFile As Long, ByVal sBuffer As String, _
'        ByVal lNumberOfBytesToRead As Long, _
'        lNumberOfBytesRead As Long) As Integer
'
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFilename As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function FtpDeleteFile Lib "wininet.dll" _
() '    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
'Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
'(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
'Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
'(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Queries an Internet option on the specified handle
'Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
'(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer
'
' Adds one or more HTTP request headers to the HTTP request handle.
'Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
'(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
'ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
'Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
'Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
'Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

'seb addition - errors and status flags
' Internet Errors
Public Const INTERNET_ERROR_BASE = 12000

Public Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Public Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Public Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Public Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Public Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Public Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Public Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Public Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Public Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Public Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Public Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Public Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Public Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Public Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Public Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Public Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Public Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Public Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Public Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Public Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Public Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Public Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Public Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Public Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Public Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Public Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Public Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Public Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Public Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Public Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Public Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Public Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)

Public Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Public Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Public Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Public Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Public Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Public Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Public Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Public Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Public Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Public Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Public Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Public Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Public Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Public Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Public Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Public Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Public Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)

' FTP API errors

Public Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110)
Public Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111)
Public Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112)

Public Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157)
Public Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158)
Public Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159)
Public Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163)
Public Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164)
Public Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165)

Public Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166)
Public Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167)
Public Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169)
Public Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170)

Public Const INTERNET_FLAG_RAW_DATA = &H40000000           ' FTP/gopher find: receive the item as raw (structured) data
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000   ' FTP: use existing InternetConnect handle for server if possible

'
' flags for InternetOpen():
'

Public Const INTERNET_FLAG_ASYNC = &H10000000              ' this request is asynchronous (where supported)

' additional flags
'

Public Const INTERNET_FLAG_SECURE = &H800000               ' use PCT/SSL if applicable (HTTP)
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000      ' use keep-alive semantics
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000     ' don't handle redirections automatically
Public Const INTERNET_FLAG_READ_PREFETCH = &H100000        ' do background read prefetch
Public Const INTERNET_FLAG_NO_COOKIES = &H80000            ' no automatic cookie handling
Public Const INTERNET_FLAG_NO_AUTH = &H40000               ' no automatic authentication handling
Public Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000     ' return cache file if net request fails


Public Const INTERNET_FLAG_NEED_FILE = &H10                ' need a file for this request

'
' flags for FTP
'

Public Const INTERNET_FLAG_TRANSFER_ASCII = FTP_TRANSFER_TYPE_ASCII       ' = &H00000001
Public Const INTERNET_FLAG_TRANSFER_BINARY = FTP_TRANSFER_TYPE_BINARY     ' = &H00000002

'
' flags field masks
'

'Public Const SECURITY_INTERNET_MASK = INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
'                                 INTERNET_FLAG_IGNORE_CERT_DATE_INVALID Or _
'                                 INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS Or _
'                                 INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP
'
'Public Const INTERNET_FLAGS_MASK = INTERNET_FLAG_RELOAD Or _
'                                 INTERNET_FLAG_RAW_DATA Or _
'                                 INTERNET_FLAG_EXISTING_CONNECT Or _
'                                 INTERNET_FLAG_ASYNC Or _
'                                 INTERNET_FLAG_PASSIVE Or _
'                                 INTERNET_FLAG_NO_CACHE_WRITE Or _
'                                 INTERNET_FLAG_MAKE_PERSISTENT Or _
'                                 INTERNET_FLAG_FROM_CACHE Or _
'                                 INTERNET_FLAG_SECURE Or _
'                                 INTERNET_FLAG_KEEP_CONNECTION Or _
'                                 INTERNET_FLAG_NO_AUTO_REDIRECT Or _
'                                 INTERNET_FLAG_READ_PREFETCH Or _
'                                 INTERNET_FLAG_NO_COOKIES Or _
'                                 INTERNET_FLAG_NO_AUTH Or _
'                                 INTERNET_FLAG_CACHE_IF_NET_FAIL Or _
'                                 SECURITY_INTERNET_MASK Or _
'                                 INTERNET_FLAG_RESYNCHRONIZE Or _
'                                 INTERNET_FLAG_HYPERLINK Or _
'                                 INTERNET_FLAG_NO_UI Or _
'                                 INTERNET_FLAG_PRAGMA_NOCACHE Or _
'                                 INTERNET_FLAG_CACHE_ASYNC Or _
'                                 INTERNET_FLAG_FORMS_SUBMIT Or _
'                                 INTERNET_FLAG_NEED_FILE Or _
'                                 INTERNET_FLAG_TRANSFER_BINARY Or _
'                                 INTERNET_FLAG_TRANSFER_ASCII
'
'
'Public Const INTERNET_OPTIONS_MASK = (Not INTERNET_FLAGS_MASK)

'
' common per-API flags (new APIs)
'

Public Const WININET_API_FLAG_ASYNC = &H1                  ' force async operation
Public Const WININET_API_FLAG_SYNC = &H4                   ' force sync operation
Public Const WININET_API_FLAG_USE_CONTEXT = &H8            ' use value supplied in dwContext (even if 0)

'
' INTERNET_NO_CALLBACK - if this value is presented as the dwContext parameter
' then no call-backs will be made for that API
'

Public Const INTERNET_NO_CALLBACK = 0


