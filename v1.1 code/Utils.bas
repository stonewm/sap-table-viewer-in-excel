Attribute VB_Name = "Utils"
Option Explicit

 Function ColltoArray(Coll As Collection) As Variant()
    Dim arr() As Variant
    Dim row As Long
    Dim col As Long
    Dim array_item As Variant
    ReDim arr(1 To Coll.Count, 1 To Coll.item(1).Count) As Variant

    row = 1
    For Each array_item In Coll
        For col = 1 To array_item.Count
            arr(row, col) = array_item(col)
        Next col
        row = row + 1
    Next
    ColltoArray = arr
End Function


Public Sub DebugPrintItab(itab As SAPTableFactoryCtrl.Table)
    Dim row As Integer
    Dim col As Integer
    
    ' Print header
    For col = 1 To itab.ColumnCount
        Debug.Print itab.ColumnName(col),
    Next
    Debug.Print
    
    For row = 1 To itab.rowcount
        For col = 1 To itab.ColumnCount
            Debug.Print itab.Value(row, col),
        Next
        ' new line
        Debug.Print
    Next
End Sub

Public Sub WriteArrayInSheet(arr() As Variant, sht As Worksheet, topLeftCell As String)
    ' Clear first
    sht.Cells.ClearContents
    
    ' write to excel
    ' 先选定一个单元格，然后根据Array的大小变化
    Dim dataRng As Range
    Set dataRng = sht.Range(topLeftCell)
    dataRng.Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
End Sub

Public Sub DebugPintArray(arr() As Variant)
    Dim i As Integer 'row
    Dim j As Integer 'col
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            Debug.Print arr(i, j),
        Next
        Debug.Print
    Next
End Sub

' 将itab转换成数组
Public Function ItabToArray(itab As SAPTableFactoryCtrl.Table) As Variant
    Dim arr() As Variant
    arr = itab.data
    
    ItabToArray = arr
End Function

Public Sub WriteData(fields() As Variant, data() As Variant, sht As Worksheet)
    '-------------------------------------------------
    ' 取消Excel的屏幕刷新和计算功能以加快速度
    '-------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 文本型在Excel显示的改善
    Dim r As Long
    Dim c As Long
    For r = 1 To UBound(data, 1) ' first dimension is row
        For c = 1 To UBound(data, 2) ' second dimension is column
            If fields(c, 4) = "C" Or fields(c, 4) = "D" Then
                data(r, c) = "'" & data(r, c)
            End If
        Next
    Next
    
    ' Clear first
    sht.Cells.ClearContents

    Dim fieldname() As Variant
    Dim fieldtext() As Variant

    Dim rowcount As Integer
    rowcount = UBound(fields, 1)
    ReDim fieldname(1 To rowcount)
    ReDim fieldtext(1 To rowcount)

    For r = 1 To UBound(fields, 1)
        fieldname(r) = fields(r, 1) ' 第一列为fieldname
        fieldtext(r) = fields(r, 5) ' 第五列为fieldtext
    Next
    
    ' top left cell
    Dim topLeftCell As Range
    Set topLeftCell = sht.Range("A1")
    
    ' fieldname和fieldtext写入工作表
    ' 第一行fieldtext
    topLeftCell.Resize(1, UBound(fieldname)).Value = fieldtext

    ' 第二行fieldtext
    topLeftCell.Offset(1, 0).Resize(1, UBound(fieldname)).Value = fieldname

    ' 从第三行开始，将splitted data写入工作表
    Dim dataRng As Range
    Set dataRng = topLeftCell.Offset(2, 0)
    dataRng.Resize(UBound(data, 1), UBound(data, 2)).Value = data
    
    '---------------------------------
    ' 恢复Excel的屏幕刷新和计算
    '---------------------------------
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Function ItemExistsInCollection(ByRef target As Variant, _
                                       ByRef container As Collection) As Boolean
    Dim candidate As Variant
    Dim found As Boolean
    
    For Each candidate In container
        Select Case True
            Case IsObject(candidate) And IsObject(target)
                found = candidate Is target
            Case IsObject(candidate), IsObject(target)
                found = False
            Case Else
                found = (candidate = target)
        End Select
        If found Then
            ItemExistsInCollection = True
            Exit Function
        End If
    Next
End Function
