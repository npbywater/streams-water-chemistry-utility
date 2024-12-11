Attribute VB_Name = "modList"
Option Compare Database
Option Explicit

'' ******************************
'' Created by: Nick Bywater
'' Created for: National Park Service, CAKN
'' License: Public Domain
'' ******************************

' *****
' DELETE items from list.
' *****
' Deletes items from a list or combo box.
Public Sub DeleteListBox(lst As Object, Optional blnDeleteValue As Boolean)
    Dim intItem As Integer

    For intItem = (lst.ListCount - 1) To 0 Step -1
        lst.RemoveItem intItem
    Next

    ' The records have been removed, but the value still exists.
    If blnDeleteValue Then
        lst.Value = Null
    End If
End Sub
