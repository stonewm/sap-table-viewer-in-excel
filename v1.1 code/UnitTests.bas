Attribute VB_Name = "UnitTests"
Option Explicit

Public Sub TestUsedRange()
    ShtSelection.UsedRange.Offset(2, 0).Select
End Sub

'Public Sub TestGetFieldsFromUI()
'    Dim fieldsColl As New Collection
'    Dim result As Boolean
'    result = GetFieldsFromUI(fieldsColl)
'
'    Dim element As Variant
'    Debug.Print fieldsColl.Count
'    For Each element In fieldsColl
'        Debug.Print element
'    Next
'End Sub

'Public Sub TestGetTableFields()
'    Call Logon
'    Dim fields As New Collection
'    Dim isSuccessful As Boolean
'    isSuccessful = GetTableFields("T030", fields)
'
'    Dim arr() As Variant
'    arr = ColltoArray(fields)
'
'    DebugPintArray arr
'End Sub


'Public Sub TestWriteFields()
'    Dim tableName As String
'    Dim fields As New Collection
'    Dim isSuccessful As Boolean
'
'    tableName = Range("table_name").Value
''    Call Logon
'
'    isSuccessful = GetTableFields(tableName, fields)
'
'    Dim fieldsArray() As Variant
'    fieldsArray = ColltoArray(fields)
'
'    Dim selectedFields() As Variant
'    ReDim selectedFields(1 To fields.Count, 1 To 2)
'    Dim r As Long
'    Dim c As Long
'    For r = 1 To fields.Count
'        selectedFields(r, 1) = fields.item(r).item(1)
'        selectedFields(r, 2) = fields.item(r).item(5)
'    Next
'
'    Call WriteArrayInSheet(selectedFields, ShtSelection, "A3")
'
'    ' 表头
'    ShtSelection.Range("A1").Value = "表名称"
'    ShtSelection.Range("table_name").Value = tableName
'    ShtSelection.Range("A2").Value = "字段名称"
'    ShtSelection.Range("B2").Value = "字段标签"
'    ShtSelection.Range("C2").Value = "显示"
'    ShtSelection.Range("D2").Value = "筛选"
'End Sub


'Public Sub TestReadTableContent()
'    Dim isSuccessful As Boolean
'    Dim data As New Collection
'    Dim fields As New Collection
'    Dim tableName As String
'    Dim filter As String
'
'    Call Logon
'
'    tableName = Range("table_name").Value
'
'    isSuccessful = GetTableFields(tableName, fields)
'    isSuccessful = ReadTableContent(tableName, data, fields, "")
'
'    Debug.Print "OK"
'End Sub

'Public Sub TestWriteTableContentInSheet()
''    Call Logon
'
'    Dim tableName As String
'    tableName = Range("table_name").Value
'
'    Dim isSuccessful As Boolean
'    Dim data As New Collection
'    Dim fields As New Collection
'    Dim filter As String
'
'    filter = GetOptions
'    isSuccessful = GetTableFields(tableName, fields)
'    isSuccessful = ReadTableContent(tableName, data, fields, filter)
'
'    Dim dataArr() As Variant
'    Dim fieldsArr() As Variant
'
'    dataArr = ColltoArray(data)
'    fieldsArr = ColltoArray(fields)
'
'    Dim sht As Worksheet
'    For Each sht In Worksheets
'        If sht.Name = tableName Then
'            Call DeleteSheet(sht)
'        End If
'    Next
'    Set sht = ThisWorkbook.Worksheets.Add
'    sht.Name = tableName
'
'    Call WriteData(fieldsArr, dataArr, sht)
'End Sub


