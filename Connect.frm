VERSION 5.00
Begin VB.Form frmConnectInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect Info"
   ClientHeight    =   2640
   ClientLeft      =   915
   ClientTop       =   1440
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   3105
   Begin VB.TextBox txtRealName 
      Height          =   288
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1500
   End
   Begin VB.TextBox txtServer 
      Height          =   288
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Co&nnect"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtPassword 
      Height          =   288
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   1500
   End
   Begin VB.TextBox txtNick 
      Height          =   288
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "Real Name"
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Nick Name"
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Server Name"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmConnectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

  Unload frmConnectInfo

End Sub


Private Sub cmdConnect_Click()

  frmConnectInfo.Hide
  
End Sub

Private Sub txtuser_Change()

End Sub


