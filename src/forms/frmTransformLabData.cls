Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'' ******************************
'' Created by: Nick Bywater
'' Created for: National Park Service, CAKN
'' Created on: 2024-January
'' License: Public Domain
'' ******************************

Enum WaterChemistryParameter
    ParameterName
    ParameterUnit
End Enum

Enum DuplicateType
    DuplicateValue
    TriplicateValue
End Enum

Private Sub cboSourceTableName_GotFocus()
    modUtility.EmptyComboBox Me.cboSourceTableName
    modUtility.ListTables Me.cboSourceTableName, SourceTable
End Sub

Private Sub cmdShowMissingParameters_Click()
    ShowMissingParameters
End Sub

Private Sub ShowMissingParameters()
    Dim db As DAO.Database
    Dim rstSource As DAO.Recordset

    Dim mapMissingParameters As Scripting.Dictionary
    Dim lngUnpivotFieldIndexes() As Long
    Dim lngMetadataEndColumn As Long
    Dim lngRowFieldCount As Long
    Dim strSourceTableName As String
    Dim varField As Variant
    Dim strSourceName As String
    Dim strCleanedName As String
    Dim blnHasComment As Boolean

    On Error GoTo ErrorHandler

    Set db = Access.CurrentDb

    strSourceTableName = Nz(Me.cboSourceTableName, "")
    lngMetadataEndColumn = Nz(Me.txtMetadataEndNum, 0)

    If strSourceTableName = "" Then
        MsgBox "Please supply a source table name.", vbInformation, "DMDB information"
        GoTo ExitSub
    ElseIf lngMetadataEndColumn <= 0 Then
        MsgBox "Please supply a metadata end number.", vbInformation, "DMDB information"
        GoTo ExitSub
    Else
        Set rstSource = db.OpenRecordset(strSourceTableName)
    End If

    lngUnpivotFieldIndexes = modDataOp.SetFieldIndexes(lngMetadataEndColumn, rstSource.Fields.Count - 1)

    lngRowFieldCount = UBound(lngUnpivotFieldIndexes) - LBound(lngUnpivotFieldIndexes) + 1

    blnHasComment = IIf(lngRowFieldCount Mod 2 = 1, True, False)

    If blnHasComment Then
        lngRowFieldCount = lngRowFieldCount - 1
    End If

    Dim i As Integer
    i = 1

    modList.DeleteListBox Me.lstMissingParameters

    For Each varField In lngUnpivotFieldIndexes()
        strSourceName = rstSource.Fields(varField).Name
        strCleanedName = CleanChemistryParameter(strSourceName)

        If i <= lngRowFieldCount Then
            If Not strSourceName Like "Date*" Then
                If Not IsWaterChemistryParameter(strCleanedName) Then
                    Me.lstMissingParameters.AddItem "SOURCE: " & strSourceName & ";" & "SOURCE cleaned: " & strCleanedName
                End If
            End If
        End If
        i = i + 1
    Next

    rstSource.Close
    Set rstSource = Nothing
    Set db = Nothing

    MsgBox "DONE!", vbInformation, "Water Chemstry Utility"

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Sub cmdTransformDataset_Click()
    Dim sfrm As Access.SubForm
    Dim strFilePath As String
    Dim strTargetTableName As String
    Dim strSourceTableName As String
    Dim lngMetadataEndColumn As Long

    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    On Error GoTo ErrorHandler

    Set db = Access.CurrentDb

    strSourceTableName = Nz(Me.cboSourceTableName, "")
    lngMetadataEndColumn = Nz(Me.txtMetadataEndNum, 0)

    If strSourceTableName = "" Then
        MsgBox "Please supply a source table name.", vbInformation, "DMDB information"
        GoTo ExitSub
    ElseIf lngMetadataEndColumn <= 0 Then
        MsgBox "Please supply a metadata end number.", vbInformation, "DMDB information"
        GoTo ExitSub
    Else
        Set rst = db.OpenRecordset(strSourceTableName)
    End If

    strTargetTableName = "stage_water_chemistry_" & modUtility.DateTimeToString(ACCESS_TYPE)

    UnpivotData rst, strTargetTableName, lngMetadataEndColumn

    If Me.chkSeparateSiteIDAndCode Then
        SeparateSiteIDAndSiteCode strTargetTableName
    End If

    Application.RefreshDatabaseWindow

    MsgBox "DONE!", vbInformation, "Water Chemstry Utility"

    rst.Close
    Set rst = Nothing
    Set db = Nothing

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical
    Resume ExitSub

End Sub

Private Sub SeparateSiteIDAndSiteCode(strTableName As String)
    Dim db As DAO.Database
    Dim strSQL As String

    Set db = Access.CurrentDb

    strSQL = "UPDATE [" & strTableName & "] " & _
            "SET SiteID = Trim$(Right$(SiteID, Len(SiteID) - InStr(SiteID,' '))), " & _
            "    SiteCode = Trim$(Left$(SiteID, InStr(SiteID,' ')))"

    db.Execute strSQL
End Sub

Private Sub UnpivotData(rstSource As DAO.Recordset, strTargetTableName As String, lngMetadataEndColumn As Long)
    Dim mapUnpivotRecordsetValues As Scripting.Dictionary
    Dim mapMetadataRecordsetValues As Scripting.Dictionary
    Dim lngMetadataIndexes() As Long
    Dim lngUnpivotFieldIndexes() As Long

    lngMetadataIndexes = modDataOp.SetFieldIndexes(1, lngMetadataEndColumn - 1)
    Set mapMetadataRecordsetValues = modDataOp.CollectRecordsetValues(rstSource, lngMetadataIndexes)

    rstSource.MoveFirst
    lngUnpivotFieldIndexes = modDataOp.SetFieldIndexes(lngMetadataEndColumn, rstSource.Fields.Count - 1)
    Set mapUnpivotRecordsetValues = modDataOp.CollectRecordsetValues(rstSource, lngUnpivotFieldIndexes)

    UnpivotToTargetTable strTargetTableName, mapMetadataRecordsetValues, mapUnpivotRecordsetValues
End Sub

Private Sub UnpivotToTargetTable(strTargetTableName As String, _
                                mapMetadataValues As Scripting.Dictionary, _
                                mapUnpivotValues As Scripting.Dictionary)
    Dim db As DAO.Database
    Dim rstTarget As DAO.Recordset
    Dim rfFieldIndex As RecordsetField
    Dim strSourceColumnName As String
    Dim strTargetTableColumnName As String
    Dim strParameterName As String
    Dim strParameterUnit As String
    Dim lngRowFieldCount As Long
    Dim colStack As VBA.Collection
    Dim rfParameter As RecordsetField
    Dim rfDate As RecordsetField
    Dim varRowIndex As Variant
    Dim colFields As VBA.Collection
    Dim varLabName As Variant
    Dim varRequesterName As Variant

    Set db = Access.CurrentDb

    modSQL.CreateTable db, strTargetTableName

    Set rstTarget = db.OpenRecordset(strTargetTableName)

    Set colStack = New VBA.Collection

    varLabName = Me.txtLabName
    varRequesterName = Me.txtRequestorName

    For Each varRowIndex In mapUnpivotValues.Keys
        Set colFields = mapUnpivotValues(varRowIndex)
        lngRowFieldCount = colFields.Count

        For Each rfFieldIndex In colFields
            Dim blnHasComment As Boolean

            blnHasComment = IIf(lngRowFieldCount Mod 2 = 1, True, False)

            If Not rfFieldIndex.FieldName Like "Date*" Then
                Push rfFieldIndex, colStack
            End If

            If rfFieldIndex.FieldName Like "Date*" Then
                Set rfParameter = Pop(colStack)
                Set rfDate = rfFieldIndex

                rstTarget.AddNew

                ' First, POPULATE metadata fields.
                rstTarget.Fields("RowNumber").Value = varRowIndex
                rstTarget.Fields("LabName").Value = varLabName
                rstTarget.Fields("RequesterName").Value = varRequesterName

                PopulateMetadataFields rstTarget, varRowIndex, mapMetadataValues

                ' Second, POPULATE parameter and measurement values.
                strSourceColumnName = CleanChemistryParameter(rfParameter.FieldName)
                strTargetTableColumnName = GetTargetColumnName(strSourceColumnName)

                strParameterName = GetWaterChemistryParameter(strSourceColumnName, ParameterName)
                strParameterUnit = GetWaterChemistryParameter(strSourceColumnName, ParameterUnit)

                rstTarget.Fields(strTargetTableColumnName).Value = strParameterName
                rstTarget.Fields("MeasurementUnit").Value = strParameterUnit
                rstTarget.Fields("Value").Value = rfParameter.FieldValue
                rstTarget.Fields("AnalysisDate").Value = rfDate.FieldValue

                If IsDuplicateColumnPrefix(strSourceColumnName) Then
                    If DuplicateColumnPrefixType(strSourceColumnName) = DuplicateValue Then
                        rstTarget.Fields("AnalysisOrder").Value = 2
                    ElseIf DuplicateColumnPrefixType(strSourceColumnName) = TriplicateValue Then
                        rstTarget.Fields("AnalysisOrder").Value = 3
                    End If
                Else
                    rstTarget.Fields("AnalysisOrder").Value = 1
                End If

                If blnHasComment Then
                    rstTarget.Fields("Notes").Value = colFields(lngRowFieldCount).FieldValue
                End If

                rstTarget.Update
            End If
        Next
    Next

    rstTarget.Close
    Set rstTarget = Nothing
    Set db = Nothing
End Sub

Private Sub PopulateMetadataFields(rst As DAO.Recordset, varRowIndex As Variant, mapMetadataValues As Scripting.Dictionary)
    Dim colRowCollection As VBA.Collection
    Dim rfFieldIndex As RecordsetField
    Dim strSourceColumnName As String
    Dim strTargetTableColumnName As String

    For Each rfFieldIndex In mapMetadataValues.Item(varRowIndex)
        strSourceColumnName = CleanChemistryParameter(rfFieldIndex.FieldName)
        strTargetTableColumnName = GetTargetColumnName(strSourceColumnName)
        rst.Fields(strTargetTableColumnName).Value = rfFieldIndex.FieldValue
    Next
End Sub

' Expects cleaned column name.
Public Function GetTargetColumnName(ByVal strColumnName As String) As String
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strParameter As String
    Dim strSQL As String

    Set db = Access.CurrentDb

    strColumnName = ReplaceDuplicateColumnPrefix(strColumnName)

    strSQL = "SELECT source_col_name, target_col_name " & _
            "FROM xref_map_column_names " & _
            "WHERE Replace(Nz(source_col_name), ' ', '') = '" & Replace(strColumnName, " ", "") & "'"

    Set rst = db.OpenRecordset(strSQL)

    If Not rst.EOF Then
        GetTargetColumnName = rst.Fields("target_col_name").Value
    Else
        Debug.Print "parameter not in xref: " & strColumnName
    End If

    rst.Close
    Set rst = Nothing
    Set db = Nothing
End Function

Public Function CleanChemistryParameter(ByVal strColumnName As String) As String
    Dim blnHasNewLine As Boolean
    Dim blnHas2Underscores As Boolean

    blnHasNewLine = modUtility.HasNewline(strColumnName)
    blnHas2Underscores = modUtility.Has2Underscores(strColumnName)

    If blnHasNewLine Then
        strColumnName = modUtility.ReplaceNewline(strColumnName, " ")
    End If

    If blnHas2Underscores Then
        strColumnName = Replace(strColumnName, "__", " ")
    End If

    CleanChemistryParameter = strColumnName

End Function

Public Function GetWaterChemistryParameter(strColumnName As String, WCParameter As WaterChemistryParameter) As Variant
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strParameter As String
    Dim strSQL As String

    Set db = Access.CurrentDb

    strParameter = modUtility.ReplaceNewline(strColumnName, " ")
    strParameter = ReplaceDuplicateColumnPrefix(strParameter)

    strSQL = "SELECT source_col_name, parameter_name, parameter_unit " & _
            "FROM xref_map_column_names " & _
            "WHERE Replace(Nz(source_col_name), ' ', '') = '" & Replace(strParameter, " ", "") & "' " & _
            "    AND target_col_name = 'WaterChemistryParameter'"

    Set rst = db.OpenRecordset(strSQL)

    If Not rst.EOF Then
        If WCParameter = ParameterName Then
            GetWaterChemistryParameter = rst.Fields("parameter_name").Value
        ElseIf WCParameter = ParameterUnit Then
            GetWaterChemistryParameter = rst.Fields("parameter_unit").Value
        End If
    End If

End Function

Public Function IsWaterChemistryParameter(strColumnName As String) As Boolean
    Dim strTargetColumnName As String

    strTargetColumnName = GetTargetColumnName(strColumnName)

    If strTargetColumnName <> "" Then
        If strTargetColumnName = "WaterChemistryParameter" Then
            IsWaterChemistryParameter = True
        End If
    End If

End Function

Private Function Replace2Underscores(strColumnName As String, strReplace As String) As String
    Replace2Underscores = Replace(strColumnName, "__", " ")
End Function

Public Function ReplaceDuplicateColumnPrefix(strColumnName As String) As String
    Select Case True
        Case strColumnName Like "Duplicate*"
            ReplaceDuplicateColumnPrefix = Trim(Replace(strColumnName, "Duplicate", ""))
        Case strColumnName Like "Triplicate*"
            ReplaceDuplicateColumnPrefix = Trim(Replace(strColumnName, "Triplicate", ""))
        Case Else
            ReplaceDuplicateColumnPrefix = strColumnName
    End Select
End Function

Public Function DuplicateColumnPrefixType(strColumnName As String) As DuplicateType
    Select Case True
        Case strColumnName Like "Duplicate*"
            DuplicateColumnPrefixType = DuplicateValue
        Case strColumnName Like "Triplicate*"
            DuplicateColumnPrefixType = TriplicateValue
    End Select
End Function

Public Function IsDuplicateColumnPrefix(strColumnName As String) As Boolean
    Select Case True
        Case strColumnName Like "Duplicate*"
            IsDuplicateColumnPrefix = True
        Case strColumnName Like "Triplicate*"
            IsDuplicateColumnPrefix = True
    End Select
End Function
