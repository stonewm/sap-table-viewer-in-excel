Attribute VB_Name = "UIInterface"
Option Explicit


Public Sub WriteTableFields()
    Dim tableName As String
    Dim fields As New Collection
    Dim isSuccessful As Boolean
    
    ' 清除3行以下数据
'    ShtSelection.rows("3:" & ShtSelection.rows.Count).ClearContents
    ShtSelection.UsedRange.Offset(2, 0).ClearContents
    
    tableName = UCase(Trim(Range("table_name").Value))
    
    If Len(tableName) = 0 Then
        MsgBox "请输入表名!", vbCritical + vbOKOnly
        Exit Sub
    End If

    isSuccessful = GetTableFields(tableName, fields)
    If isSuccessful = False Then Exit Sub
    
    Dim fieldsArray() As Variant
    fieldsArray = ColltoArray(fields)
    
    Dim selectedFields() As Variant
    ReDim selectedFields(1 To fields.Count, 1 To 3)
    Dim r As Long
    Dim c As Long
    For r = 1 To fields.Count
        selectedFields(r, 1) = fields.item(r).item(1)
        selectedFields(r, 2) = fields.item(r).item(5)
        selectedFields(r, 3) = "X"
    Next

    Call WriteArrayInSheet(selectedFields, ShtSelection, "A3")

    ' 表头
    ShtSelection.Range("A1").Value = "表名称"
    ShtSelection.Range("table_name").Value = tableName
    ShtSelection.Range("A2").Value = "字段名称"
    ShtSelection.Range("B2").Value = "字段标签"
    ShtSelection.Range("C2").Value = "显示"
    ShtSelection.Range("D2").Value = "筛选"
End Sub


Public Sub WriteTableContentInSheet()
    Dim tableName As String
    tableName = UCase(Trim(Range("table_name").Value))
    
    Dim isSuccessful As Boolean
    Dim data As New Collection
    Dim fields As New Collection
    Dim requiredFields As New Collection
    Dim targetFields As New Collection
    Dim filter As String
    
    If Len(tableName) = 0 Then
        MsgBox "请输入表名!", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    filter = GetOptions
    isSuccessful = GetTableFields(tableName, fields)
    If isSuccessful = False Then Exit Sub
    
    isSuccessful = GetFieldsFromUI(requiredFields)
    
    Dim element As Variant
    Dim tempColl As Collection
    For Each element In fields
        If ItemExistsInCollection(element.item(1), requiredFields) Then
            Set tempColl = New Collection
            tempColl.Add item:=element.item(1), key:="FIELDNAME"
            tempColl.Add item:=element.item(2), key:="OFFSET"
            tempColl.Add item:=element.item(3), key:="LENGTH"
            tempColl.Add item:=element.item(4), key:="TYPE"
            tempColl.Add item:=element.item(5), key:="FIELDTEXT"

            targetFields.Add tempColl
        End If
    Next
    
    isSuccessful = ReadTableContent(tableName, data, targetFields, filter)
    If isSuccessful = False Then Exit Sub
      
    Dim dataArr() As Variant
    Dim fieldsArr() As Variant
    
    dataArr = ColltoArray(data)
    fieldsArr = ColltoArray(targetFields)
    
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Name = tableName Then
            Call DeleteSheet(sht)
        End If
    Next
    Set sht = ThisWorkbook.Worksheets.Add
    sht.Name = tableName
        
    Call WriteData(fieldsArr, dataArr, sht)
End Sub

Public Sub DeleteSheet(sht As Worksheet)
    Application.DisplayAlerts = False
    sht.Delete
    Application.DisplayAlerts = True
End Sub

Public Function GetFieldsFromUI(fieldsColl As Collection) As Boolean
    GetFieldsFromUI = False
    
    Dim i As Integer
    For i = 3 To 2000
        If ShtSelection.Range("A" & i).Value = "" Then Exit For
        If ShtSelection.Range("C" & i).Value = "X" Then
            fieldsColl.Add item:=ShtSelection.Range("A" & i).Value
        End If
    Next
End Function

Public Function GetOptions()
    GetOptions = ""
    
    Dim sht As Worksheet
    Set sht = ShtSelection
    
    Dim i As Integer
    i = 3
    Dim criteria As String
    Dim isFirst As Boolean
    isFirst = False
    
    For i = 3 To 2000
        If sht.Range("A" & i).Value = "" Then Exit For
        
        If sht.Range("D" & i).Value <> "" Then
            If criteria = "" Then
                isFirst = True
            End If
                                     
            If isFirst Then
                criteria = sht.Range("A" & i).Value + _
                           " = '" + CStr(sht.Range("D" & i).Value) + "'"
                isFirst = False
            Else
                criteria = criteria + " AND " + _
                           sht.Range("A" & i).Value + " = '" + _
                           CStr(sht.Range("D" & i).Value) + "'"
            End If
        End If
    Next
    
    GetOptions = criteria
End Function




