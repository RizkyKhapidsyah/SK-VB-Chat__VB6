VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   3000
   ClientLeft      =   1155
   ClientTop       =   2070
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3000
   ScaleWidth      =   3150
   Begin VB.CheckBox chkChanFlag 
      Caption         =   "&MIC Only"
      Height          =   252
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   972
   End
   Begin VB.CheckBox chkChanFlag 
      Caption         =   "&Local"
      Height          =   252
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      Width           =   972
   End
   Begin VB.OptionButton opChanType 
      Caption         =   "Pro&tected"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.OptionButton opChanType 
      Caption         =   "Pu&blic"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton opChanType 
      Caption         =   "Pri&vate"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtMinUsers 
      Height          =   288
      Left            =   1800
      TabIndex        =   6
      Text            =   "1"
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtMaxUsers 
      Height          =   288
      Left            =   1800
      TabIndex        =   7
      Text            =   "25"
      Top             =   2040
      Width           =   500
   End
   Begin VB.TextBox txtSearch 
      Height          =   288
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1812
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   972
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&List"
      Default         =   -1  'True
      Height          =   372
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Minimum Users"
      Height          =   252
      Left            =   360
      TabIndex        =   12
      Top             =   1680
      Width           =   1404
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum Users"
      Height          =   252
      Left            =   360
      TabIndex        =   11
      Top             =   2040
      Width           =   1404
   End
   Begin VB.Label Label8 
      Caption         =   "Search for"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lChanTypes As Long
Dim lChanFlags As Long
Dim lOpName As Long

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

  Unload frmSearch
  
End Sub

Private Sub cmdList_Click()
  Dim szSearchString As String
  Dim iMinUsers As Integer
  Dim iMaxUsers As Integer
  Dim i As Integer
  
  
  szSearchString = txtSearch.text
  lChanTypes = 0
  For i = 0 To 2
    opChanType_Click (i)
  Next i
  lChanFlags = 0
  For i = 0 To 1
    chkChanFlag_Click (i)
  Next i
  lOpName = opContains
  iMinUsers = CLng(txtMinUsers.text)
  iMaxUsers = CLng(txtMaxUsers.text)
  
  Unload frmSearch
  frmChanList.Show
  ClearList frmChanList, frmChanList.lstChannels
  
  g_fChanList = True
  If Not (g_Socket.ListAllChannels(iMinUsers, iMaxUsers, szSearchString, lOpName, lChanTypes, lChanFlags)) Then
    MsgBox "Unable to list requested channels", "List Channel Error", vbOKOnly
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


