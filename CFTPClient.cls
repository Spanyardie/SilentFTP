VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFTPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCommandDate As Date
Private mvarCommandTime As Date
Private mvarHostIP As String
Private mvarHostPort As Long
Private mvarSessionID As String

Private mvarCommandEvents As CFTPCommandEvents

Private mvarFilePath As String
Private mvarFTPHostName As String
Private mvarFTPUserName As String
Private mvarFTPPassword As String

Private Const MODULE As String = "CFTPClient::"

Public Event StateChanged(ByVal StateString As String)

Public Sub SendFTP(ByVal FileName As String)

    Dim sSep As String
    Dim sFileNameAndPath As String
    Dim lRet As Long
    
    On Error GoTo SendFTP_Error
    
    sSep = ""
    
    If Trim$(mvarFilePath) = "" Then
        sFileNameAndPath = App.Path & "\" & FileName & ".dat"
    Else
        If InStr(Trim$(mvarFilePath), "\") = 0 Then
            sSep = "\"
        End If
        sFileNameAndPath = Trim$(mvarFilePath) & sSep & Trim$(FileName) & ".dat"
    End If
        
    lRet = FTPSendFile(sFileNameAndPath, Trim$(FileName & ".dat"))
    If lRet <> 0 Then
        'error sending ftp file
        'this is the only bit we care about, fail miserably
        Err.Raise lRet, MODULE & "SendFTP", "Failed sending file by FTP."
    End If
        
Exit_Properly:
    Exit Sub
    
SendFTP_Error:
    Err.Raise Err.Number, MODULE & "SendFTP - " & Err.Source, Err.Description
    
End Sub

Public Property Let FTPPassword(ByVal vData As String)
    
    gsFTPPassword = Trim$(vData)
    
End Property

Public Property Get FTPPassword() As String
    
    FTPPassword = gsFTPPassword
    
End Property

Public Property Let FTPUserName(ByVal vData As String)
    
    gsFTPUsername = Trim$(vData)
    
End Property

Public Property Get FTPUserName() As String
    
    FTPUserName = gsFTPUsername
    
End Property

Public Property Let FTPHostName(ByVal vData As String)
    
    gsFTPHostName = Trim$(vData)
    
End Property

Public Property Get FTPHostName() As String
    
    FTPHostName = gsFTPHostName
    
End Property

Public Property Let FilePath(ByVal vData As String)
    
    mvarFilePath = Trim(vData)
    
End Property

Public Property Get FilePath() As String
    
    FilePath = mvarFilePath
    
End Property

Public Sub InitFTPClient()

    Dim lRet As Long
    
    'open internet
    lRet = FTPInternetOpen()
    
    If lRet = 0 Then
        'couldn't do it, raise an error
        Err.Raise vbObjectError + 2000, MODULE & "InitFTPClient - ", "Could not open Internet Connection"
    End If
    
    'set up status callback function
    glStatusCallback = InternetSetStatusCallback(ghOpen, AddressOf FTPStatusCallback)
    
    If glStatusCallback = INTERNET_INVALID_STATUS_CALLBACK Then
        Err.Raise vbObjectError + 2001, MODULE & "InitFTPClient", "Unable to set Internet Status Callback"
    End If
        
    'callback is active
    gbStatusCallbackActive = True
    
    'try to connect
    lRet = FTPConnect()
    If lRet = 0 Then
        'couldn't create a session over internet, bomb out
        If gbStatusCallbackActive Then
            InternetSetStatusCallback ghOpen, 0
            gbStatusCallbackActive = False
        End If
        FTPInternetClose
        Err.Raise vbObjectError + 2002, MODULE & "InitFTPClient", "Unable to create FTP Session"
    End If
    
    'all ok
    
End Sub

Public Property Get CommandEvents() As CFTPCommandEvents
    
    Set CommandEvents = mvarCommandEvents
    
End Property

Public Property Let HostPort(ByVal vData As Long)
    
    mvarHostPort = vData
    
End Property

Public Property Get HostPort() As Long
    
    HostPort = mvarHostPort
    
End Property

Public Property Let HostIP(ByVal vData As String)
    
    mvarHostIP = vData
    
End Property

Public Property Get HostIP() As String
    
    HostIP = mvarHostIP
    
End Property

Public Property Let CommandTime(ByVal vData As Date)
    
    mvarCommandTime = vData
    
End Property

Public Property Get CommandTime() As Date
    
    CommandTime = mvarCommandTime
    
End Property

Public Property Let CommandDate(ByVal vData As Date)
    
    mvarCommandDate = vData
    
End Property

Public Property Get CommandDate() As Date
    
    CommandDate = mvarCommandDate
    
End Property

Private Sub Class_Initialize()

    Set mvarCommandEvents = New CFTPCommandEvents
    
    'global holder
    Set gpFTPClient = Me
    
    gbFTPFileSent = False
    
End Sub

Private Sub Class_Terminate()

    Set mvarCommandEvents = Nothing
    
    Set gpFTPClient = Nothing
    
End Sub

Public Sub StateChanged(ByVal State As String)

    RaiseEvent StateChanged(State)
    
End Sub

Private Function FTPInternetOpen() As Long

    ghOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

    FTPInternetOpen = ghOpen
    
End Function

Private Function FTPInternetClose() As Long

    If ghConnection <> 0 Then
        InternetCloseHandle (ghConnection)
    End If
    
    'do we have a status callback active
    If gbStatusCallbackActive Then
        'only if internet is open
        If ghOpen <> 0 Then
            InternetSetStatusCallback ghOpen, 0
            gbStatusCallbackActive = False
        End If
    End If
    
    If ghOpen <> 0 Then
        InternetCloseHandle (ghOpen)
    End If
    
    ghConnection = 0
    ghOpen = 0
    
    gbActiveSession = False

End Function

Private Function FTPConnect() As Long

    Dim tContext As FTP_FILE_PUT_CONTEXT
    
    With tContext
        .sType = "Connect"
        .sFilename = ""
        .sServerIP = mvarHostIP
        .sServerPort = mvarHostPort
    End With
    
    ghConnection = InternetConnect(ghOpen, gsFTPHostName, INTERNET_INVALID_PORT_NUMBER, _
     gsFTPUsername, gsFTPPassword, INTERNET_SERVICE_FTP, INTERNET_FLAG_ACTIVE, VarPtr(tContext))
     
    If ghConnection = 0 Then
        gbActiveSession = False
    Else
        gbActiveSession = True
    End If

    FTPConnect = ghConnection
    
End Function

Private Function FTPDisconnect() As Long

    If ghConnection <> 0 Then
        InternetCloseHandle ghConnection
    End If
    
    ghConnection = 0
    
    gbActiveSession = False

End Function

Private Function FTPSendFile(ByVal LocalFile As String, ByVal RemoteFile As String) As Long

    Dim bRet As Boolean
    Dim tContext As FTP_FILE_PUT_CONTEXT
    
    FTPSendFile = 0
    
    With tContext
        .sType = "FTPPutFile"
        .sFilename = RemoteFile
        .sServerIP = mvarHostIP
        .sServerPort = mvarHostPort
    End With
    
    bRet = FtpPutFile(ghConnection, LocalFile, RemoteFile, _
     FTP_TRANSFER_TYPE_BINARY, VarPtr(tContext))

    If bRet = False Then
        'error occured putting file, return error status
        FTPSendFile = Err.LastDllError
    End If
       
End Function

Public Property Get FTPFileSent() As Boolean

    FTPFileSent = gbFTPFileSent
    
End Property

Public Sub CloseFTPClient()

    'disconnect and close
    FTPDisconnect
    
    FTPInternetClose
    
    'we done
    
End Sub

'seb addition - directory stuff
Public Function GetRemoteFolderList() As Boolean

End Function

Public Function ChangeToFolder(ByVal sFolderName As String) As Boolean
    
End Function

Public Function CreateFolder(ByVal sFolderName As String) As Boolean

End Function

Public Function DeleteFolder(ByVal sFolderName As String) As Boolean

End Function

Public Function GetFileList(ByVal sFolderName As String) As Boolean

End Function

