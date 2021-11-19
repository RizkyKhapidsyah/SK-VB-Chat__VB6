VERSION 5.00
Begin VB.Form frmChanList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Channel List"
   ClientHeight    =   3675
   ClientLeft      =   1230
   ClientTop       =   1530
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   3105
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   2160
      TabIndex        =   2
      Top             =   3240
      Width           =   732
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate"
      Height          =   372
      Left            =   1200
      TabIndex        =   1
      Top             =   3240
      Width           =   732
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "&Join"
      Default         =   -1  'True
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   732
   End
   Begin VB.ListBox lstChannels 
      Columns         =   1
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Channel Name"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3972
   End
End
Attribute VB_Name = "frmChanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()

  Unload frmChanList
  
End Sub

Private Sub cmdCreate_Click()
  
  If g_Channel.Valid Then
    g_Channel.Leave (False)
    If IsDebug Then Debug.Print "Channel.Leave issued"
  End If
  
  frmCreateChan.Show 1
  frmCreateChan.cmdJoin.Visible = False
  Unload frmChanList
  
  
End Sub


Private Sub cmdJoin_Click()
Dim szChanName As String
Dim szPassword As String
Dim iUserFlags As Integer
Dim i As Integer

  iUserFlags = 0
  szPassword = ""

  If Not (Selected(frmChanList.lstChannels)) Then
    Exit Sub
  End If
  
  With frmChanList
    With lstChannels
      For i = 0 To .ListCount - 1
        If .Selected(i) Then
           szChanName = .List(i)
            Exit For
        End If
      Next i
    End With
  End With
    
  If g_Channel.Valid Then
    g_Channel.Leave (False)
    If IsDebug Then Debug.Print "Channel.Leave issued"
  End If
  
  If g_Socket.JoinChannel(szChanName, szPassword, iUserFlags) Then
    g_fChannel = False
    Unload frmChanList
  End If
  
End Sub


