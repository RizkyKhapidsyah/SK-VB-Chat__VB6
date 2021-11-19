VERSION 5.00
Begin VB.Form frmJoinChan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Join Channel"
   ClientHeight    =   1095
   ClientLeft      =   360
   ClientTop       =   1650
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1095
   ScaleWidth      =   4650
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox txtChanName 
      Height          =   288
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1332
   End
   Begin VB.TextBox txtPassword 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1332
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "&Join"
      Default         =   -1  'True
      Height          =   372
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Channel Name"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1812
   End
End
Attribute VB_Name = "frmJoinChan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdJoin_Click()
  
  If g_Channel.Valid Then
    g_Channel.Leave (False)
    If IsDebug Then Debug.Print "Channel.Leave issued"
  End If

  If g_Socket.JoinChannel(frmJoinChan.txtChanName.text, frmJoinChan.txtPassword.text) Then
    g_fChannel = False
    Unload frmJoinChan
    Unload frmChanList
  End If

End Sub


Private Sub Form_Load()
Dim i As Integer

  With frmChanList
    With .lstChannels
      For i = 0 To .ListCount - 1
        If .Selected(i) Then
           frmJoinChan.txtChanName.text = .ItemData(i)
            Exit For
        End If
      Next i
    End With
  End With

End Sub


