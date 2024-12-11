Attribute VB_Name = "modStack"
Option Compare Database
Option Explicit

'' ******************************
'' Created by: Nick Bywater
'' Created for: National Park Service, CAKN
'' License: Public Domain
'' ******************************

Public Sub Push(var As Variant, col As VBA.Collection)
    col.Add var
End Sub

Public Function Pop(col As VBA.Collection) As Variant
    Dim var As Variant
    Dim lngCount As Long

    lngCount = col.Count

    If lngCount > 0 Then
        If IsObject(col.Item(lngCount)) Then
            Set var = col.Item(lngCount)
        Else
            var = col.Item(lngCount)
        End If
        col.Remove lngCount

        If IsObject(var) Then
            Set Pop = var
        Else
            Pop = var
        End If
    Else
        Err.Raise vbObjectError + 1030, , "Stack is empty"
    End If
End Function

Public Function IsEmpty(col As VBA.Collection) As Boolean
    If col.Count = 0 Then
        IsEmpty = True
    End If
End Function

