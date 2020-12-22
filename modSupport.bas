Attribute VB_Name = "modSupport"
Option Explicit

Public gbActiveSession As Boolean
Public ghOpen As Long, ghConnection As Long
Public gdwType As Long
Public glStatusCallback As Long
Public gbStatusCallbackActive As Boolean

Public gsFTPHostName As String
Public gsFTPUsername As String
Public gsFTPPassword As String

Public gbFTPFileSent As Boolean

Public gpFTPClient As CFTPClient

Public Sub Main()

    'initialise here
    gbActiveSession = False
    ghOpen = 0
    ghConnection = 0
    
End Sub

Public Sub FTPStatusCallback(ByVal hInternetSession As Long, ByVal nContext As Long, _
ByVal lStatus As Long, ByVal StatusInfo As Long, ByVal StatusInfoLength As Long)

    Dim sState As String
    
    sState = ""
    
    Select Case lStatus
        Case INTERNET_STATUS_RESOLVING_NAME:
            'resolving host
            sState = "Resolving Host '" & gsFTPHostName & "'"
            
        Case INTERNET_STATUS_NAME_RESOLVED:
            'host resolved
            sState = "Host '" & gsFTPHostName & "' resolved"
            
        Case INTERNET_STATUS_CONNECTING_TO_SERVER:
            'connecting
            sState = "Connecting to Host"
            
        Case INTERNET_STATUS_CONNECTED_TO_SERVER:
            'connected
            sState = "Connected successfully"
                        
        Case INTERNET_STATUS_CLOSING_CONNECTION:
            'disconnecting
            sState = "Disconnecting from '" & gsFTPHostName & "'..."
            
        Case INTERNET_STATUS_CONNECTION_CLOSED:
            'disconnected
            sState = "Disconnected"
            
        Case INTERNET_STATUS_REQUEST_COMPLETE:
            'response completed
            sState = "Request completed"
                        
    End Select
       
    'route back thru ftpclient to raise event to client application
    If sState <> "" Then
        gpFTPClient.StateChanged sState
    End If
    
End Sub
