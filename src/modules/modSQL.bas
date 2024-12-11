Attribute VB_Name = "modSQL"
Option Compare Database
Option Explicit

' ******************************
' Created by: Nick Bywater
' Created for: National Park Service, CAKN
' Application created on: 2020-June
' License: Public Domain
' ******************************

' CreateTable
Public Sub CreateTable(dbTarget As DAO.Database, strTableName As String)
    Dim db As DAO.Database
    Dim tblNewTable As DAO.TableDef
    Dim rstXref As DAO.Recordset
    Dim strSQL As String
    Dim strPrefixName As String
    Dim strTColumnName As String
    Dim strTColumnType As String
    Dim strTColumnSize As String

    Set tblNewTable = dbTarget.CreateTableDef(strTableName)

    Set db = Access.CurrentDb

    strTColumnName = "target_column_name"
    strTColumnType = "column_type"
    strTColumnSize = "column_size"

    strSQL = "SELECT * " & _
        "FROM sys_schema_xref WHERE " & strTColumnName & " IS NOT NULL " & _
        "ORDER BY target_column_order"

    Set rstXref = db.OpenRecordset(strSQL)

    ' Create empty table.
    While Not rstXref.EOF
        Dim strTargetTableColumnName As String
        Dim intFldType As Integer
        Dim lngFldSize As Long

        strTargetTableColumnName = rstXref.Fields(strTColumnName).Value
        intFldType = rstXref.Fields(strTColumnType).Value
        lngFldSize = rstXref.Fields(strTColumnSize).Value

        Dim fldNew As DAO.Field
        Set fldNew = tblNewTable.CreateField(strTargetTableColumnName, intFldType, lngFldSize)

        tblNewTable.Fields.Append fldNew

        rstXref.MoveNext

        Set fldNew = Nothing
    Wend

    dbTarget.TableDefs.Append tblNewTable

    Set tblNewTable = Nothing
    rstXref.Close
    Set rstXref = Nothing
    Set db = Nothing

End Sub

