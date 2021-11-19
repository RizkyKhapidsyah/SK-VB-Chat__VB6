VERSION 5.00
Begin VB.Form frmChatVB 
   ClientHeight    =   5985
   ClientLeft      =   5970
   ClientTop       =   1500
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5985
   ScaleWidth      =   5775
   Visible         =   0   'False
   Begin VB.ListBox lstMembers 
      Height          =   4935
      Left            =   4560
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdWhisper 
      Caption         =   "&Whisper"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtMsg 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   4455
   End
   Begin VB.TextBox txtHistory 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   4455
   End
   Begin VB.PictureBox EventSink 
      Height          =   2415
      Left            =   3120
      ScaleHeight     =   2355
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Menu mnuChan 
      Caption         =   "&Channels"
      Begin VB.Menu mnuChanListChan 
         Caption         =   "List Channels"
      End
      Begin VB.Menu mnuChanCreate 
         Caption         =   "C&reate"
      End
      Begin VB.Menu mnuChanJoin 
         Caption         =   "&Join"
      End
      Begin VB.Menu mnuChanExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMembers 
      Caption         =   "&Members"
      Begin VB.Menu mnuMemRealName 
         Caption         =   "Get &Real Name"
      End
      Begin VB.Menu mnuMemHost 
         Caption         =   "Make &Host"
      End
      Begin VB.Menu mnuMemSpeaker 
         Caption         =   "Make &Speaker"
      End
      Begin VB.Menu mnuMemSpectator 
         Caption         =   "Make S&pectator"
      End
      Begin VB.Menu mnuMemNoWhisper 
         Caption         =   "No &Whispers"
      End
   End
End
Attribute VB_Name = "frmChatVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
  
  InsertText txtMsg.text, g_Channel.GetMe.GetName
  g_Channel.SendText txtMsg.text
  txtMsg.text = ""
  
End Sub

Private Sub cmdWhisper_Click()
Dim i As Integer
Dim lPointer As Long
Dim lCount As Long
Dim rgmem() As Object
Dim msg As String

  msg = CStr(frmChatVB.txtMsg.text)
  frmChatVB.txtMsg.text = ""
  
  With lstMembers
    If Not (.SelCount = 0) Then
      ReDim rgmem(0 To (.SelCount - 1)) As Object
      lCount = 0
      For i = 0 To .ListCount - 1
        If ((lCount >= 10) Or (lCount > .SelCount)) Then
          Exit For
        End If
        If .Selected(i) Then
          lPointer = .ItemData(i)
          Set rgmem(lCount) = g_Chatsock.ObjFromPointer(lPointer)
          lCount = lCount + 1
        End If
      Next i
    End If
  End With
  If Not g_Channel.SendTextList(msg, rgmem, lCount) Then
    MsgBox "Unable to send message to recepients", vbOKOnly, "Send Whisper Error"
  End If
  
End Sub


Private Sub EventSink_ChannelAddMember(ByVal member As Object, ByVal eventID As Long)
If IsDebug Then Debug.Print "Channel Add Member"

  With lstMembers
    .AddItem member.GetName
    .ItemData(.NewIndex) = member
    Set member = Nothing
  End With
  
  EventSink.DoneWithEvent (eventID)
  
End Sub

Private Sub EventSink_ChannelDelMember(ByVal member As Object, ByVal eventID As Long)
Dim i As Integer

  With lstMembers
    For i = 0 To .ListCount - 1
      If .ItemData(i) = member Then
        .RemoveItem (i)
        .Refresh
        Exit For
      End If
    Next i
  End With
  EventSink.DoneWithEvent (eventID)
  
End Sub


Private Sub EventSink_ChannelGotMemlist(ByVal eventID As Long)

If IsDebug Then Debug.Print "Channel Got MemList"

  EventSink.DoneWithEvent (eventID)
End Sub

Private Sub EventSink_ChannelModeMember(ByVal member As Object, ByVal modePrevious As Long, ByVal eventID As Long)
Dim lModeVal As Long

  lModeVal = member.GetMemberMode
  If lModeVal = membermodeNoMemUpdates Then
    GoTo Done
  Else
    Select Case lModeVal
      Case membermodeSpectator
        InsertText "is now a Spectator", member.GetName()
      Case membermodeHost, (membermodeHost + membermodeSpeaker)
        InsertText "is now a Host", member.GetName()
      Case membermodeSpeaker
        InsertText "is now a Speaker", member.GetName()
      Case membermodeNoWhisper
        InsertText "does not accept Whispers", member.GetName()
      Case Else
        MsgBox "The membermode " & CStr(lModeVal) & " is not handled in this SDK sample", vbOKOnly, "MemberMode"
    End Select
    Set member = Nothing
  End If

Done:
  EventSink.DoneWithEvent (eventID)
  
End Sub


Private Sub EventSink_ChannelText(ByVal text As String, ByVal member As Object, ByVal eventID As Long)

If IsDebug Then Debug.Print "Channel Text"

  InsertText text, member.GetName()
  Set member = Nothing
  EventSink.DoneWithEvent (eventID)

End Sub


Private Sub EventSink_ChannelWhisperText(ByVal szText As String, ByVal memberFrom As Object, ByVal membersTo As Object, ByVal eventID As Long)
Dim objMem As Object

If IsDebug Then Debug.Print "Whisper Text"

  For Each objMem In membersTo
    MsgBox szText, vbOKOnly, "Whispered Message from " & memberFrom.GetName() & " to " & objMem.GetName()
  Next objMem
  Set objMem = Nothing
  
  EventSink.DoneWithEvent (eventID)
   
End Sub


Private Sub EventSink_ChatSockError(ByVal error As Long, ByVal eventID As Long)
'  g_fChanGetProp = Not (g_fChanGetProp)
'  g_fChanList = Not (g_fChanList)
'  g_fMemList = Not (g_fMemList)
'  g_fMemGetProp = Not (g_fMemGetProp)

    Select Case error
    Case errorServerFull
        MsgBox "ERROR:" + CStr(error) + "  errorServerFull"
    Case errorAuthOnly
        MsgBox "ERROR:" + CStr(error) + "  errorAuthOnly"
    Case errorUnknownSecPackage
        MsgBox "ERROR:" + CStr(error) + "  errorUnknownSecPackage"
    Case errorNotIrc
        MsgBox "ERROR:" + CStr(error) + "  errorNotIrc"
    Case errorIrcNotAllowed
        MsgBox "ERROR:" + CStr(error) + "  errorIrcNotAllowed"
    Case errorBadUsername
        MsgBox "ERROR:" + CStr(error) + "  errorBadUsername"
    Case errorBadNickname
        MsgBox "ERROR:" + CStr(error) + "  errorBadNickname"
    Case errorBadChannelGuid
        MsgBox "ERROR:" + CStr(error) + "  errorBadChannelGuid"
    Case errorAlreadyLoggedIn
        MsgBox "ERROR:" + CStr(error) + "  errorAlreadyLoggedIn"
    Case errorNoSuchNick
        MsgBox "ERROR:" + CStr(error) + "  errorNoSuchNick"
    Case errorInvalidRecipList
        MsgBox "ERROR:" + CStr(error) + "  errorInvalidRecipList"
    Case errorNotModerated
        MsgBox "ERROR:" + CStr(error) + "  errorNotModerated"
    Case errorNickCollision
        MsgBox "ERROR:" + CStr(error) + "  errorNickCollision"
    Case errorBadChannelName
        MsgBox "ERROR:" + CStr(error) + "  errorBadChannelName"
    Case errorNoWhisper
        MsgBox "ERROR:" + CStr(error) + "  errorNoWhisper"
    Case errorInvalidPassword
        MsgBox "ERROR:" + CStr(error) + "  errorInvalidPassword"
        frmJoinChan.Show 1
    Case errorAuthNotAvail
        MsgBox "ERROR:" + CStr(error) + "  errorAuthNotAvail"
    Case errorBanned
        MsgBox "ERROR:" + CStr(error) + "  errorBanned"
    Case errorNoMatches
        MsgBox "ERROR:" + CStr(error) + "  errorNoMatches"
    Case errorPropLookup
        MsgBox "ERROR:" + CStr(error) + "  errorPropLookup"
    Case errorNotOperator
        MsgBox "ERROR:" + CStr(error) + "  errorNotOperator"
    Case errorNotMemberProp
        MsgBox "ERROR:" + CStr(error) + "  errorNotMemberProp"
    Case errorNotChannelProp
        MsgBox "ERROR:" + CStr(error) + "  errorNotChannelprop"
    Case errorTooManyTerms
        MsgBox "ERROR:" + CStr(error) + "  errorTooManyTerms"
    Case errorTooManyProp
        MsgBox "ERROR:" + CStr(error) + "  errorTooManyProp"
    Case errorNoMoreProp
        MsgBox "ERROR:" + CStr(error) + "  errorNoMoreProp"
    Case errorProperty
        MsgBox "ERROR:" + CStr(error) + "  errorProperty"
    Case errorPropMode
        MsgBox "ERROR:" + CStr(error) + "  errorPropMode"
    Case errorUnicodeNotAllowed
        MsgBox "ERROR:" + CStr(error) + "  errorUnicodeNotAllowed"
    Case errorNotSpeaker
        MsgBox "ERROR:" + CStr(error) + "  errorNotSpeaker"
    Case errorNotYou
        MsgBox "ERROR:" + CStr(error) + "  errorNotYou"
    Case errorIsHost
        MsgBox "ERROR:" + CStr(error) + "  errorIsHost"
    Case errorCallerNotHost
        MsgBox "ERROR:" + CStr(error) + "  errorCallerNotHost"
    Case errorNotMIC
        MsgBox "ERROR:" + CStr(error) + "  errorNotMIC"
    Case errorService
        MsgBox "ERROR:" + CStr(error) + "  errorService"
    Case errorSecurity
        MsgBox "ERROR:" + CStr(error) + "  errorSecurity"
    Case errorServer
        MsgBox "ERROR:" + CStr(error) + "  errorServer"
    Case errorByteCount
        MsgBox "ERROR:" + CStr(error) + "  errorByteCount"
    Case errorChannelBadPass
        MsgBox "ERROR:" + CStr(error) + "  errorChannelBadPass"
    Case errorInviteOnlyChannel
        MsgBox "ERROR:" + CStr(error) + "  errorInviteOnlyChannel"
    Case errorTooManyChannels
        MsgBox "ERROR:" + CStr(error) + "  errorTooManyChannels"
    Case errorNotInChannel
        MsgBox "ERROR:" + CStr(error) + "  errorNotInChannel"
    Case errorAlreadyOnChannel
        MsgBox "ERROR:" + CStr(error) + "  errorAlreadyOnChannel"
    Case errorChannelFull
        MsgBox "ERROR:" + CStr(error) + "  errorChannelFull"
    Case errorCantMakeUniqueChan
        MsgBox "ERROR:" + CStr(error) + "  errorCantMakeUniqueChan"
    Case errorChannelNotFound
        MsgBox "ERROR:" + CStr(error) + "  errorChannelNotFound"
    Case errorChannelExists
        MsgBox "ERROR:" + CStr(error) + "  errorChannelExists"
    Case errorCancelFail
        MsgBox "ERROR:" + CStr(error) + "  errorCancelFail"
    Case errorJoinFail
        MsgBox "ERROR:" + CStr(error) + "  errorJoinFail"
    Case errorCreateFail
        MsgBox "ERROR:" + CStr(error) + "  errorCreateFail"
    Case errorChannelCancel
        MsgBox "ERROR:" + CStr(error) + "  errorChannelCancel"
    Case errorClose
        MsgBox "ERROR:" + CStr(error) + "  errorClose"
    Case errorIllegalUser
        MsgBox "ERROR:" + CStr(error) + "  errorIllegalUser"
    Case errorAliasInUse
        MsgBox "ERROR:" + CStr(error) + "  errorAliasInUse"
    Case errorUnknownUser
        MsgBox "ERROR:" + CStr(error) + "  errorUnknownUser"
    Case errorNotLoggedIn
        MsgBox "ERROR:" + CStr(error) + "  errorNotLoggedIn"
    Case errorHostDropped
        MsgBox "ERROR:" + CStr(error) + "  errorHostDropped"
    Case errorNetworkDown
        MsgBox "ERROR:" + CStr(error) + "  errorNetworkDown"
    Case errorSocketClosed
        MsgBox "ERROR:" + CStr(error) + "  errorSocketClosed"
    Case errorLostConnection
        MsgBox "ERROR:" + CStr(error) + "  errorLostConnection"
    Case errorInvalidSocket
        MsgBox "ERROR:" + CStr(error) + "  errorInvalidSocket"
    Case errorSocketError
        MsgBox "ERROR:" + CStr(error) + "  errorSocketError"
    Case errorNoData
        MsgBox "ERROR:" + CStr(error) + "  errorNoData"
    Case errorTimeout
        MsgBox "ERROR:" + CStr(error) + "  errorTimeout"
    Case errorCantSend
        MsgBox "ERROR:" + CStr(error) + "  errorCantSend"
    Case errorCantConnect
        MsgBox "ERROR:" + CStr(error) + "  errorCantConnect"
    Case errorSocketCreate
        MsgBox "ERROR:" + CStr(error) + "  errorSocketCreate"
    Case errorHostNotFound
            MsgBox "ERROR:" + CStr(error) + "  errorHostNotFound"
    Case errorWinsockDll
        MsgBox "ERROR:" + CStr(error) + "  errorWinsockDll"
    Case errorNotConnected
        MsgBox "ERROR:" + CStr(error) + "  errorNotConnected"
    Case errorQueueEmpty
        MsgBox "ERROR:" + CStr(error) + "  errorQueueEmpty"
    Case errorNotInList
        MsgBox "ERROR:" + CStr(error) + "  errorNotInList"
    Case errorAlreadyInList
        MsgBox "ERROR:" + CStr(error) + "  errorAlreadyInList"
    Case errorFirstChar
        MsgBox "ERROR:" + CStr(error) + "  errorFirstChar"
    Case errorIllegalChars
        MsgBox "ERROR:" + CStr(error) + "  errorIllegalChars"
    Case errorTooMuchData
        MsgBox "ERROR:" + CStr(error) + "  errorTooMuchData"
    Case errorNotUnicode
        MsgBox "ERROR:" + CStr(error) + "  errorNotUnicode"
    Case errorNotAnsi
        MsgBox "ERROR:" + CStr(error) + "  errorNotAnsi"
    Case errorExiting
        MsgBox "ERROR:" + CStr(error) + "  errorExiting"
    Case errorStringTooLong
        MsgBox "ERROR:" + CStr(error) + "  errorStringTooLong"
    Case errorEvent
        MsgBox "ERROR:" + CStr(error) + "  errorEvent"
    Case errorWait
        MsgBox "ERROR:" + CStr(error) + "  errorWait"
    Case errorVersion
        MsgBox "ERROR:" + CStr(error) + "  errorVersion"
    Case errorSystem
         MsgBox "ERROR:" + CStr(error) + "  errorSystem"
    Case Else
      MsgBox "UNKOWN ERROR - " + CStr(error)
    End Select
   
    EventSink.DoneWithEvent (eventID)

End Sub

Private Sub EventSink_PropertyData(ByVal fLastRecord As Boolean, ByVal fNestedParent As Boolean, ByVal property As Object, ByVal eventID As Long)
Dim i As Integer
Dim iPropCount As Integer
Dim objPropData As Object

If IsDebug Then Debug.Print "Property Data"

' List All Channels data
  If g_fChanList Then
    Set objPropData = property.GetProperty(indexName)
    With objPropData
      If .fString And .fAnsi Then
        frmChanList.lstChannels.AddItem .szData
      End If
    End With
    If fLastRecord Then
      g_fChanList = False
    End If
    GoTo Done
  End If

' GetRealName data
  If g_fGetRealName Then
    Set objPropData = property.GetProperty(indexName)
    With objPropData
      If .fString And .fAnsi Then
        MsgBox "Member's real name is " + .szData, vbOKOnly, "Real Name"
      End If
    End With
    If fLastRecord Then
      g_fChanList = False
    End If
    GoTo Done
  End If
  
    
'
Done:
  EventSink.DoneWithEvent (eventID)
  
End Sub


Private Sub EventSink_SocketAddChannel(ByVal channel As Object, ByVal eventID As Long)
If IsDebug Then Debug.Print "Socket Add Channel"
  
  If Not (g_fChannel) Then
    Set g_Channel = channel
    frmChatVB.Visible = True
    frmChatVB.EventSink.MonitorChannel g_Channel

If IsDebug Then Debug.Print "Channel Created and Monitoring"
    
    Set channel = Nothing
    g_fChannel = True
    frmChatVB.Caption = g_Channel.GetName() & " discussing " & g_Channel.GetTopic()
  Else
    g_fChannel = False
  End If
  EventSink.DoneWithEvent (eventID)
  
End Sub


Private Sub EventSink_SocketDelChannel(ByVal channel As Object, ByVal eventID As Long)
If IsDebug Then Debug.Print "Socket DEL CHANNEL"

  g_fChannel = True
  If ChanCleanup(channel, g_fChannel) Then
    Set channel = Nothing
    g_fChannel = False
  End If
  
  EventSink.DoneWithEvent (eventID)
  
End Sub

Private Sub Form_Load()

  If Not ConnectSrv Then
    DisconnectSrv
    frmChatVB.Hide
    End
  End If
   
End Sub


Private Sub mnuChanCreate_Click()
  
  If g_Channel.Valid Then
    g_Channel.Leave (False)
    If IsDebug Then Debug.Print "Channel.Leave issued"
  End If
  
  frmCreateChan.Show
  With frmCreateChan
    .cmdJoin.Visible = False
    .cmdCreate.Visible = True
    .chkChanFlag(0).Visible = True
    .chkChanFlag(1).Visible = True
    .opChanType(0).Visible = True
    .opChanType(1).Visible = True
    .opChanType(2).Visible = True
    .txtMaxUsers.Enabled = True
    .txtChanTopic.Enabled = True
  End With

End Sub

Private Sub mnuChanExit_Click()

  DisconnectSrv
  Unload frmChatVB
  End
End Sub

Private Sub mnuChanJoin_Click()
  
  frmCreateChan.Show
  With frmCreateChan
    .cmdJoin.Visible = True
    .cmdCreate.Visible = False
    .chkChanFlag(0).Visible = False
    .chkChanFlag(1).Visible = False
    .opChanType(0).Visible = False
    .opChanType(1).Visible = False
    .opChanType(2).Visible = False
    .txtMaxUsers.Enabled = False
    .txtChanTopic.Enabled = False
  End With
 
End Sub






Private Sub mnuChanListChan_Click()

  frmSearch.Show 1

End Sub

Private Sub mnuMemHost_Click()
Dim objMem As Object
Dim fChanModeMem As Boolean

  If Not (Selected(frmChatVB.lstMembers)) Then
    Exit Sub
  Else
    fChanModeMem = True
    Set objMem = GetMemObj(frmChatVB.lstMembers)
    If Not (objMem.MakeHost) Then
      MsgBox "Unable to make " & objMem.GetName() _
              & " a Host", vbOKOnly, "Make Host"
      fChanModeMem = False
    End If
  End If
End Sub

Private Sub mnuMemNoWhisper_Click()
Dim objMem As Object
Dim fOn As Boolean

  If Not (Selected(frmChatVB.lstMembers)) Then
    Exit Sub
  Else
  ' toggle context
    mnuMemNoWhisper.Checked = Not (mnuMemNoWhisper.Checked)
    fOn = mnuMemNoWhisper.Checked
    
    g_fChanModeMem = True
    Set objMem = GetMemObj(frmChatVB.lstMembers)
    If Not (objMem.SetNoWhisper(fOn)) Then
      MsgBox "Unable to make to set No Whispers for " & objMem.GetName(), _
                                            vbOKOnly, "Set No Whispers"
      mnuMemNoWhisper.Checked = Not (mnuMemNoWhisper.Checked)
      g_fChanModeMem = False
    End If
  End If
  
End Sub


Private Sub mnuMemRealName_Click()
Dim objMem As Object
Dim fGetRealName As Boolean

  If Not (Selected(frmChatVB.lstMembers)) Then
    Exit Sub
  Else
    g_fGetRealName = True
    Set objMem = GetMemObj(frmChatVB.lstMembers)
    If Not (objMem.GetRealName) Then
      MsgBox "Unable to return Real Name of " & lstMembers.ItemData(lstMembers.ListIndex), "Get Real Name", vbOKOnly
      g_fGetRealName = False
    End If
  End If
End Sub


Private Sub mnuMemSpeaker_Click()
Dim objMem As Object
'Dim fChanModeMem As Boolean

  If Not (Selected(frmChatVB.lstMembers)) Then
    Exit Sub
  Else
    g_fChanModeMem = True
    Set objMem = GetMemObj(frmChatVB.lstMembers)
    If Not (objMem.MakeSpeaker) Then
      MsgBox "Unable to make " & lstMembers.ItemData(lstMembers.ListIndex) & " a Speaker", "Make Speaker", vbOKOnly
      g_fChanModeMem = False
    End If
  End If
  
End Sub


Private Sub mnuMemSpectator_Click()
Dim objMem As Object
'Dim fChanModeMem As Boolean

  If Not (Selected(frmChatVB.lstMembers)) Then
    Exit Sub
  Else
    g_fChanModeMem = True
    Set objMem = GetMemObj(frmChatVB.lstMembers)
    If Not (objMem.MakeSpectator) Then
      MsgBox "Unable to make " & lstMembers.ItemData(lstMembers.ListIndex) & " a Spectator", "Make Spectator", vbOKOnly
      g_fChanModeMem = False
    End If
  End If
End Sub



