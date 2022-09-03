Attribute VB_Name = "SAPTableReadModule"
Option Explicit


'-----------------------------------------------------------
' 获取SAP table的所有字段，存储在fieldArray中(variant数组)
'-----------------------------------------------------------
Public Function GetTableFields(tableName As String, fieldsArray() As Variant) As Boolean
    Dim functions As New SAPFunctionsOCX.SAPFunctions
    Dim fm As SAPFunctionsOCX.Function
    Dim fieldsTable As SAPTableFactoryCtrl.Table
    
    GetTableFields = False
    
    If sapConnection Is Nothing Or sapConnection.IsConnected <> tloRfcConnected Then
        Debug.Print "Please connect to SAP."
        Exit Function
    End If
    
    Set functions.Connection = sapConnection
    Set fm = functions.Add("RFC_READ_TABLE")
        
    '----------------------------
    '填充Importing parameters
    '----------------------------
    fm.Exports("QUERY_TABLE").Value = tableName  'QUERY_TABLE: Table name
    fm.Exports("DELIMITER").Value = "~"          'DELIMITER是输出时字段的分割符
    fm.Exports("NO_DATA").Value = "X"

    Set fieldsTable = fm.Tables("FIELDS")        'FIELDS表示要输出的列
    
    fm.Call
    
    '如果有Exception,说明有错误产生
    If fm.Exception <> "" Then
        Debug.Print fm.Exception
        Exit Function
    End If

    fieldsArray = Utils.ItabToArray(fieldsTable)
        
    GetTableFields = True
End Function


'-----------------------------------------------------------
' 获取SAP table的data，存储在variant数组中
'-----------------------------------------------------------
Public Function GetTableContent(tableName As String, data() As Variant, fields() As Variant, filter As String) As Boolean
    Dim functions As New SAPFunctionsOCX.SAPFunctions
    Dim fm As SAPFunctionsOCX.Function
    
    ' RFC_READ_TABLE的三个table型参数
    Dim optionsTable As SAPTableFactoryCtrl.Table
    Dim dataTable As SAPTableFactoryCtrl.Table
    Dim fieldsTable As SAPTableFactoryCtrl.Table
    
    GetTableContent = False
    
    If sapConnection Is Nothing Or sapConnection.IsConnected <> tloRfcConnected Then
        Debug.Print "Please connect to SAP."
        Exit Function
    End If
    
    Set functions.Connection = sapConnection
    Set fm = functions.Add("RFC_READ_TABLE")
    Dim delimiter As String
    delimiter = "~"
    
    '填充Importing parameters
    fm.Exports("QUERY_TABLE").Value = tableName   'QUERY_TABLE是要查找的表名
    fm.Exports("DELIMITER").Value = delimiter     'DELIMITER是输出时字段的分割符

    'table参数
    Set optionsTable = fm.Tables("OPTIONS")  'OPTIONS是筛选条件
    Set fieldsTable = fm.Tables("FIELDS")    'FIELDS表示要输出的列
    Set dataTable = fm.Tables("DATA")        'DATA为输出的数据
    
    If filter <> "" Then
        optionsTable.rows.Add
        optionsTable.Value(1, "TEXT") = filter
    End If
    
    fm.Call
    
    '如果有Exception,说明有错误产生
    If fm.Exception <> "" Then
        Debug.Print fm.Exception
        Exit Function
    End If
    
    ' 存储fields信息的数组
    fields = ItabToArray(fieldsTable)
    
    ' 存储data信息的数组
    Dim rawData() As Variant
    rawData = ItabToArray(dataTable)
    
    ' 将data分割
    data = splitData(rawData, delimiter)
    
    GetTableContent = True
End Function


Private Function splitData(data() As Variant, delimeter As String) As Variant
    Dim dataSplitted() As Variant '返回值
    Dim rowcount As Long
    Dim colcount As Long
    
    rowcount = UBound(data, 1)
   
    ' 列数需要计算
    Dim testcol As Variant
    testcol = Split(data(1, 1), delimeter) '根据第一个数据来确定列数

    colcount = UBound(testcol) + 1
    ReDim dataSplitted(1 To rowcount, 1 To colcount)
    
    Dim line As Variant
    Dim r As Long
    Dim c As Long
    For r = 1 To rowcount
        line = Split(data(r, 1), delimeter) ' line从0开始
        For c = 1 To colcount
            dataSplitted(r, c) = line(c - 1)
        Next
    Next
    
    splitData = dataSplitted
End Function




