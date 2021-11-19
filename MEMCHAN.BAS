Attribute VB_Name = "MemChan"
Option Explicit


Function ClearList(frmForm As Form, cControl As Control) As Boolean
Dim i As Integer

' clear listbox of previous info
  With frmForm
    With cControl
    If .ListCount >= 1 Then
      For i = (.ListCount - 1) To 0 Step -1
      .RemoveItem i
      Next i
    End If
    End With
  End With
End Function

Function GetMemObj(lstControl As Control) As Object
' returns member object

  Dim lPointer As Long
  
  With lstControl
    lPointer = .ItemData(.ListIndex)
    Set GetMemObj = g_Chatsock.ObjFromPointer(lPointer)
    End With
    
End Function


Public Sub InsertText(text As String, membername As String)
Dim newtext As String

  newtext = frmChatVB.txtHistory.text
  newtext = newtext & membername & ": " & vbCrLf
  newtext = newtext & "    " & text
  newtext = newtext & vbCrLf
  frmChatVB.txtHistory.text = newtext
End Sub

Function Selected(lstControl As Control) As Boolean
' checks if a selection is made in a list control

  With lstControl
    If Not (.SelCount > 0) Then
      MsgBox "Must select a member", vbOKOnly, "No Selection"
    Else
      Selected = True
    End If
  End With
End Function

Function VerifyChan(szChan As String, fmic As Boolean, floc As Boolean) As Boolean

  If g_Chatsock.VerifyCreateChannelName(szChan, fmic, floc) Then
    VerifyChan = True
  Else
    VerifyChan = False
  End If
End Function


