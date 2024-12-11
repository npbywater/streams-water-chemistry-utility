Attribute VB_Name = "modDataOp"
Option Compare Database
Option Explicit

'' ******************************
'' Created by: Nick Bywater
'' Created for: National Park Service, CAKN
'' Created on: 2024-January
'' License: Public Domain
'' ******************************

' The recordset 'rst' must have a field named 'RowID'.
' This field's values MUST be a unique identifier for each
' recordset record in 'rst'.

' This function returns a dictionary of field collections.
Public Function CollectRecordsetValues(rst As DAO.Recordset, lngFieldIndexes() As Long, Optional allowNullValues As Boolean) As Scripting.Dictionary
    Dim RowID As Long
    Dim mapRecordsetFields As Scripting.Dictionary
    Set mapRecordsetFields = New Scripting.Dictionary

    While Not rst.EOF
        RowID = rst.Fields("RowID").Value

        mapRecordsetFields.Add RowID, CollectRecordValue(RowID, rst, lngFieldIndexes)

        rst.MoveNext
    Wend

    Set CollectRecordsetValues = mapRecordsetFields

End Function

Public Function CollectRecordValue(RowID As Long, rst As DAO.Recordset, lngFieldIndexes() As Long, Optional allowNullValues As Boolean) As VBA.Collection
    Dim fld As DAO.Field
    Dim varIndex As Variant
    Dim varValue As Variant

    Dim colRecordsetFields As VBA.Collection

    Set colRecordsetFields = New VBA.Collection

    For Each varIndex In lngFieldIndexes()
        Dim rfld As RecordsetField
        Set rfld = New RecordsetField

        varValue = Nz(Trim(rst.Fields(varIndex).Value), "")

        If varValue <> "" And Not allowNullValues Then
            rfld.RecordID = RowID
            rfld.FieldName = rst.Fields(varIndex).Name
            rfld.FieldValue = varValue

            colRecordsetFields.Add rfld
        ElseIf allowNullValues Then
            rfld.RecordID = RowID
            rfld.FieldName = rst.Fields(varIndex).Name
            rfld.FieldValue = Null

            colRecordsetFields.Add rfld
        End If

        Set rfld = Nothing
    Next

    Set CollectRecordValue = colRecordsetFields

End Function

' This returns an array of field indexes for required recordset fields.
Public Function SetFieldIndexes(lngStart As Long, lngEnd As Long) As Long()
    Dim lngFieldIndexes() As Long
    Dim i As Long
    Dim j As Long

    j = 0

    ReDim lngFieldIndexes(0 To (lngEnd - lngStart))

    For i = lngStart To lngEnd
        lngFieldIndexes(j) = i
        j = j + 1
    Next

    SetFieldIndexes = lngFieldIndexes

End Function
