Attribute VB_Name = "SAPTableReadModule"
Option Explicit


'-----------------------------------------------------------
' 获取SAP table的所有字段，存储在fieldArray中(variant数组)
'-----------------------------------------------------------
Public Function GetTableFields(tableName As String, fieldsColl As Collection) As Boolean
    Dim functions As New SAPFunctionsOCX.SAPFunctions
    Dim fm As SAPFunctionsOCX.Function
    Dim fieldsTable As SAPTableFactoryCtrl.Table
    
    GetTableFields = False
    
    If sapConnection Is Nothing Then
        MsgBox "请登录SAP系统!", vbOKOnly + vbInformation
        Exit Function
    End If
    If sapConnection.IsConnected <> tloRfcConnected Then
        MsgBox "请登录SAP系统!", vbOKOnly + vbInformation
        Exit Function
    End If
    
    Set functions.Connection = sapConnection
    Set fm = functions.Add("RFC_READ_TABLE")
        
    '填充Importing parameters
    fm.Exports("QUERY_TABLE").Value = tableName  'QUERY_TABLE: Table name
    fm.Exports("DELIMITER").Value = "~"          'DELIMITER是输出时字段的分割符
    fm.Exports("NO_DATA").Value = "X"

    Set fieldsTable = fm.Tables("FIELDS")        'FIELDS表示要输出的列
    
    fm.Call
    
    '如果有Exception,说明有错误产生
    If fm.Exception <> "" Then
        MsgBox fm.Exception, vbOKOnly + vbCritical
        Exit Function
    End If
    
    Dim row As SAPTableFactoryCtrl.row
    Dim fld As Collection
    Dim requiredFields As New Collection
    For Each row In fieldsTable.rows
        Set fld = New Collection
        fld.Add item:=row.Value("FIELDNAME"), key:="FIELDNAME"
        fld.Add item:=row.Value("OFFSET"), key:="OFFSET"
        fld.Add item:=row.Value("LENGTH"), key:="LENGTH"
        fld.Add item:=row.Value("TYPE"), key:="TYPE"
        fld.Add item:=row.Value("FIELDTEXT"), key:="FIELDTEXT"
        fieldsColl.Add fld
    Next row
        
    GetTableFields = True
End Function


'--------------------------------------------------------
' 调用RFC_READ_TABLE, 调用结果存放在dataColl中
'--------------------------------------------------------
Private Function CallRfcReadTable(oFunction As SAPFunctionsOCX.Function, _
                                  oFieldsTable As SAPTableFactoryCtrl.Table, _
                                  oDataColl As Collection) As Boolean
    CallRfcReadTable = False

    Dim row         As Collection
    Dim oDataTable  As SAPTableFactoryCtrl.Table
    Dim oDataRow    As Object
    Dim oField      As Object

    '调用函数
    oFunction.Call
    If oFunction.Exception <> "" Then
        MsgBox ("调用RFC_READ_TABLE出现错误:" & oFunction.Exception)
        Exit Function
    End If

    Set oDataTable = oFunction.Tables("DATA") ' DATA参数
    
    '如果是第一次调用
    If oDataColl.Count = 0 Then
        For Each oDataRow In oDataTable.rows
            Set row = New Collection
            For Each oField In oFieldsTable.rows ' oFieldTable每一行代表一列
                row.Add item:=Trim(Mid(oDataRow.Value("WA"), _
                                   CInt(oField.Value("OFFSET")) + 1, _
                                   CInt(oField.Value("LENGTH")))), _
                        key:=oField.Value("FIELDNAME")
            Next oField
            
            oDataColl.Add item:=row, key:=CStr(oDataColl.Count)
        Next oDataRow
        
    Else ' 不是第一次调用，将DATA添加到oDataColl
        For Each oDataRow In oDataTable.rows
            Set row = oDataColl.item(CStr(oDataRow.index - 1))
        
            For Each oField In oFieldsTable.rows
                row.Add item:=Trim(Mid(oDataRow.Value("WA"), _
                                   CInt(oField.Value("OFFSET")) + 1, _
                                   CInt(oField.Value("LENGTH")))), _
                        key:=oField.Value("FIELDNAME")
            Next oField
        Next oDataRow
    End If

    oDataTable.FreeTable

    CallRfcReadTable = True
End Function


Public Function ReadTableContent(tableName As String, _
                                 oDataColl As Collection, _
                                 oFields As Collection, _
                                 filter As String) As Boolean
    'Set return value
    ReadTableContent = False
    
    Dim functions As New SAPFunctionsOCX.SAPFunctions
    Dim oFunction As SAPFunctionsOCX.Function 'SAP Function module object
    Dim row  As Collection                    'Table row object
    Dim oFieldsTable As SAPTableFactoryCtrl.Table
    Dim optionsTable As SAPTableFactoryCtrl.Table

    Set functions.Connection = sapConnection
    Set oFunction = functions.Add("RFC_READ_TABLE")
    
    'Set function module parameters
    oFunction.Exports("QUERY_TABLE") = tableName
    oFunction.Exports("DELIMITER") = ""
    oFunction.Exports("NO_DATA") = ""
    oFunction.Exports("ROWSKIPS") = 0
    oFunction.Exports("ROWCOUNT") = 50000

    Set oFieldsTable = oFunction.Tables("FIELDS")
    Set optionsTable = oFunction.Tables("OPTIONS")
    
    ' options
    If filter <> "" Then
        optionsTable.rows.Add
        optionsTable.Value(1, "TEXT") = filter
    End If
    
    'Initialize
    oFieldsTable.FreeTable
    Dim length As Integer
    length = 0

    For Each row In oFields
        'Calls every 512 bytes (The limit of RFC_READ_TABLE)
        length = length + row.item("LENGTH")
        If length <= 512 Then
            oFieldsTable.rows.Add
            oFieldsTable(oFieldsTable.rowcount, "FIELDNAME") = row.item("FIELDNAME")
        Else
            If CallRfcReadTable(oFunction, oFieldsTable, oDataColl) = False Then
                Exit Function
            End If
            oFieldsTable.FreeTable
            length = row.item("LENGTH")
            oFieldsTable.rows.Add
            oFieldsTable(oFieldsTable.rowcount, "FIELDNAME") = row.item("FIELDNAME")
        End If
    Next row

    If CallRfcReadTable(oFunction, oFieldsTable, oDataColl) = False Then
        Exit Function
    End If

    ReadTableContent = True
End Function




