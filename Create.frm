VERSION 5.00
Begin VB.Form frmCreateChan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Channel"
   ClientHeight    =   2820
   ClientLeft      =   750
   ClientTop       =   1500
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   Begin VB.CheckBox chkChanFlag 
      Caption         =   "&MIC Only"
      Height          =   252
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   2400
      Value           =   1  'Checked
      Width           =   972
   End
   Begin VB.CheckBox chkChanFlag 
      Caption         =   "&Local"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox txtMaxUsers 
      Height          =   288
      Left            =   3720
      TabIndex        =   9
      Text            =   "25"
      Top             =   2364
      Width           =   612
   End
   Begin VB.OptionButton opChanType 
      Caption         =   "Pu&blic"
      Height          =   372
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      Value           =   -1  'True
      Width           =   1000
   End
   Begin VB.OptionButton opChanType 
      Caption         =   "Pro&tected"
      Height          =   372
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   1000
   End
   Begin VB.OptionButton opChanType 
      Caption         =   "Pri&vate"
      Height          =   372
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "&Join"
      Height          =   372
      Left            =   3600
      TabIndex        =   11
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   3600
      TabIndex        =   10
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate"
      Default         =   -1  'True
      Height          =   372
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox txtPassword 
      Height          =   288
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1332
   End
   Begin VB.TextBox txtChanTopic 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1332
   End
   Begin VB.TextBox txtChanName 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1332
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label4 
      Caption         =   "Max Users"
      Height          =   252
      Left            =   2640
      TabIndex        =   15
      Top             =   2400
      Width           =   972
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   4440
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   372
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1812
   End
   Begin VB.Label Label2 
      Caption         =   "Channel Topic"
      Height          =   372
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1812
   End
   Begin VB.Label Label1 
      Caption         =   "Channel Name"
      Height          =   372
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1812
   End
End
Attribute VB_Name = "frmCreateChan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iUserMax As Integer
Dim lChanTypes As Long
Dim lChanFlags As Long
Dim lChanCreate As Long
Dim lUserFlags As Long




Private Sub chkChanFlag_Click(Index As Integer)
' returns the ChanFlag value
  Select Case Index
    Case 0
      If chkChanFlag(Index).Value = 1 Then
        lChanFlags = lChanFlags + flagLocal
      End If
    Case 1
      If chkChanFlag(Index).Value = 1 Then
        lChanFlags = lChanFlags + flagMicOnly
      End If
    Case Else
      MsgBox " Invalid Channel Type Set", "Channel Type Error", vbOKOnly
  End Select


End Sub

Private Sub cmdCancel_Click()

  Unload frmCreateChan
  ' no channel already exists then disconnect from server
  If Not (g_fChannel) Then
    DisconnectSrv
  End If
      
End Sub


Private Sub cmdCreate_Click()
' Create/Join channel

Dim szChanName As String
Dim szChanTopic As String
Dim szPassword As String
Dim i As Integer

  With frmCreateChan
    szChanName = .txtChanName.text
    If Len(szChanTopic) <> 0 Then
      szChanTopic = "<Nothing Specific>"
    Else
      szChanTopic = .txtChanTopic.text
    End If
    szPassword = .txtPassword.text
    iUserMax = CInt(.txtMaxUsers.text)
    
    lChanTypes = 0
    For i = 0 To 2
      opChanType_Click (i)
    Next i
    lChanFlags = 0
    For i = 0 To 1
      chkChanFlag_Click (i)
    Next i
    lChanCreate = createflagJoin
    lUserFlags = 0

Verify:                       ' Is Channel name valid
    If VerifyChan(szChanName, chkChanFlag(1).Value, chkChanFlag(0).Value) Then
      g_fChannel = False
      ' Create specified channel
      If g_Socket.CreateChannel(lChanTypes, lChanFlags, lChanCreate, lUserFlags, _
                                                  szChanName, szChanTopic, szPassword, iUserMax) Then
       g_fChannel = False
        Unload frmCreateChan
        Exit Sub
        ' Channel already exists Join instead
        ElseIf g_Socket.JoinChannel(szChanName, szPassword) Then
          g_fChannel = False
          Unload frmCreateChan
          Exit Sub
      End If
    Else
      MsgBox "Channel Name invalid. Please enter new name.", "Channel Name Conflict", vbOKOnly
      frmCreateChan.SetFocus
      GoTo Verify
    End If
  End With
  DisconnectSrv

End Sub


Private Sub cmdJoin_Click()
Dim szChanName As String
Dim szPassword As String

  g_fChannel = False
  With frmCreateChan
    szChanName = .txtChanName.text
    szPassword = .txtPassword.text
  End With
  If g_Socket.JoinChannel(szChanName, szPassword, 0) Then
    Unload frmCreateChan
    g_fChannel = True
  Else
    MsgBox "Unable to join the chat - " & CStr(szChanName), "Channel Join Failure", vbOKOnly
    Unload frmCreateChan
    frmCreateChan.SetFocus
    cmdJoin.Visible = False
    cmdCreate.Visible = True
  End If
End Sub


Private Sub opChanType_Click(Index As Integer)
' returns the ChannelType value
  Select Case Index
    Case 0
      If opChanType(Index).Value = True Then
        lChanTypes = lChanTypes + typePrivate
      End If
    Case 1
      If opChanType(Index).Value = True Then
        lChanTypes = lChanTypes + typePublic
      End If
    Case 2
      If opChanType(Index).Value = True Then
        lChanTypes = lChanTypes + typeProtected
      End If
    Case Else
      MsgBox " Invalid Channel Type Set", "Channel Type Error", vbOKOnly
  End Select


End Sub


