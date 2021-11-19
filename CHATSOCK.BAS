Attribute VB_Name = "ChatSock"
Option Explicit

Public Const IsDebug = True

'Global objects
Public g_Chatsock As Object
Public g_Factory As Object
Public g_Socket As Object
Public g_Channel As Object
Public g_Member As Object


' Global flags
Public g_fChatsock As Boolean
Public g_fSocket As Boolean
Public g_fChannel As Boolean
Public g_fChanList As Boolean
Public g_fGetRealName As Boolean
Public g_fChanModeMem As Boolean



Function ChanCleanup(channel As Object, fChanClose) As Boolean
' closes a channel releasing all references
  If fChanClose Then
    channel.Leave (False)
    Set channel = Nothing
    ChanCleanup = True
  End If
  
End Function

Public Sub ChatSockCleanup(ChatSock As Object, fChatSocketClose)
' Releases references to ChatSocket

  If fChatSocketClose Then
    Set ChatSock = Nothing
    fChatSocketClose = False
  End If
End Sub

Public Function ConnectSrv()

' create a chatsock object
  Set g_Chatsock = CreateObject("MIC.Chatsock")
  If Not (g_Chatsock.Valid) Then
    g_fChatsock = False
    Exit Function
  Else
    g_fChatsock = True
  End If
  
' create a chatsocket factory object
  Dim ChatsockCurrentVersion
  Set g_Factory = g_Chatsock.CreateChatSocketFactory(ChatsockCurrentVersion)
  If Not (g_Factory.Valid) Then
    ConnectSrv = False
    Exit Function
  End If
  
  frmConnectInfo.Show 1
  
' create a socket
  Set g_Socket = g_Factory.MakeSocket(frmConnectInfo.txtServer.text)
  If Not g_Socket.Valid Then
    ConnectSrv = False
    Exit Function
  Else
    g_fSocket = True
  End If
  
' login to socket
' server must allow anonymous logins, no usermode, no security checking
  g_Socket.Login True, 0, frmConnectInfo.txtRealName.text, frmConnectInfo.txtPassword.text, frmConnectInfo.txtNick.text, "", ""
  
  If Not (g_Socket.WaitUntilLoggedOn) Then
    MsgBox "Unable to log onto server " & frmConnectInfo.txtServer.text & ". Please try again later", vbOKOnly, "Login Failure"
    ConnectSrv = False
    Unload frmChatVB
    Exit Function
  Else
    If IsDebug Then Debug.Print "LOGGED ON"
  End If
  
' start to monitor the socket for messages
  If Not (frmChatVB.EventSink.MonitorSocket(g_Socket)) Then
    MsgBox "Unable to monitor socket message", vbOKOnly, "Socket Monitoring Failure"
    ConnectSrv = False
    Exit Function
  Else
If IsDebug Then Debug.Print "Monitoring Socket"
  End If
  ConnectSrv = True
  frmCreateChan.Show 1
  
End Function








Public Sub DisconnectSrv()

If g_fChannel Then
  ChanCleanup g_Channel, g_fChannel
End If

If g_fSocket Then
  SocketCleanup g_Socket, g_fSocket
End If

If g_fChatsock Then
  ChatSockCleanup g_Chatsock, g_fChatsock
End If
  

End Sub

Public Sub SocketCleanup(socket As Object, fSockClose)
' Logs off of socket and releases references
  
  If fSockClose Then
    socket.LogOff
    socket.CloseSocket
    Set socket = Nothing
    fSockClose = False
  End If
    
End Sub

