Attribute VB_Name = "modUtility"
Option Compare Database
Option Explicit

'' ******************************
'' Created by: Nick Bywater
'' Created for: National Park Service, CAKN
'' License: Public Domain
'' ******************************

Public Enum NameFormat
    FILE_TYPE
    ACCESS_TYPE
End Enum

Public Function DateTimeToString(nf As NameFormat) As String
    Dim dt As Date
    Dim strDate As String
    Dim strTime As String

    dt = Now()

    strDate = Format(dt, "YYYY-MM-DD")

    If nf = FILE_TYPE Then
        strTime = Format(dt, "HH.NN.SS")
    Else: nf = ACCESS_TYPE
        strTime = Format(dt, "HH:NN:SS")
    End If

    DateTimeToString = strDate & "T" & strTime
End Function

Public Function HasNewline(var As Variant) As Boolean
    If InStr(var, vbCrLf) > 0 Or _
        InStr(var, vbCr) > 0 Or _
        InStr(var, vbLf) Then
        HasNewline = True
    End If
End Function

Public Function Has2Underscores(var As Variant) As Boolean
    If InStr(1, var, "__") > 0 Then
        Has2Underscores = True
    End If
End Function

' ReplaceNewline
Public Function ReplaceNewline(var As Variant, Optional strReplace As String) As Variant
    ' Need to check 'vbcrlf' first, since 'cr' and 'lf' are contained in it.
    If InStr(var, vbCrLf) > 0 Then
        ReplaceNewline = Replace(var, vbCrLf, strReplace)
'        Debug.Print "vbcrlf: " & ReplaceNewline
    ElseIf InStr(var, vbCr) > 0 Then
        ReplaceNewline = Replace(var, vbCr, strReplace)
'        Debug.Print "vbcr: " & ReplaceNewline
    ElseIf InStr(var, vbLf) > 0 Then
        ReplaceNewline = Replace(var, vbLf, strReplace)
'        Debug.Print "vblf: " & ReplaceNewline
    Else
        ReplaceNewline = var
'        Debug.Print "nothing: " & ReplaceNewline
    End If
End Function

Public Sub ListTables(cboTarget As Access.ComboBox, Optional tType As TableType = AllTables)
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef

    Set db = Access.CurrentDb

    For Each tbl In db.TableDefs
        If tbl.Attributes = 0 Then
            If tType = AllTables Then
                cboTarget.AddItem tbl.Name
            ElseIf tType = SourceTable Then
                If tbl.Name Like "source_*" Then
                    cboTarget.AddItem tbl.Name
                End If
            ElseIf tType = TargetTable Then
                If tbl.Name Like "target_*" Then
                    cboTarget.AddItem tbl.Name
                End If
            End If
        End If
    Next
End Sub

Public Sub EmptyComboBox(cbo As Access.ComboBox)
    Dim x As Long

    For x = cbo.ListCount - 1 To 0 Step -1
        cbo.RemoveItem x
    Next
End Sub
