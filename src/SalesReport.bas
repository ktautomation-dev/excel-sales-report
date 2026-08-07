Attribute VB_Name = "Module1"
Option Explicit

Sub 売上を集計する()
  
    On Error GoTo ErrorHandler

    Dim errorRow As Long
    Dim hasError As Boolean

    Dim wsData As Worksheet
    Dim wsStaff As Worksheet
    Dim wsProduct As Worksheet

    Dim lastRow As Long
    
    Dim staffLastRow As Long
    Dim productLastRow As Long
    Dim i As Long

    Dim staffDict As Object
    Dim productDict As Object

    Dim staffName As String
    Dim productName As String
    Dim salesAmount As Double

    Dim key As Variant
    Dim outputRow As Long

    Set wsData = ThisWorkbook.Worksheets("売上データ")
    Set wsStaff = ThisWorkbook.Worksheets("担当者別集計")
    Set wsProduct = ThisWorkbook.Worksheets("商品別集計")

    Set staffDict = CreateObject("Scripting.Dictionary")
    Set productDict = CreateObject("Scripting.Dictionary")

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "売上データがありません。", vbExclamation
        Exit Sub
    End If
    
    hasError = False
    
    For i = 2 To lastRow
    
        If Trim(wsData.Cells(i, "A").Value) = "" _
           Or Trim(wsData.Cells(i, "B").Value) = "" _
           Or Trim(wsData.Cells(i, "C").Value) = "" _
           Or Trim(wsData.Cells(i, "D").Value) = "" _
           Or Trim(wsData.Cells(i, "E").Value) = "" Then
    
            errorRow = i
            hasError = True
            Exit For
    
        End If
    
        If Not IsNumeric(wsData.Cells(i, "D").Value) _
           Or Not IsNumeric(wsData.Cells(i, "E").Value) Then
    
            errorRow = i
            hasError = True
            Exit For
    
        End If
    
        If CDbl(wsData.Cells(i, "D").Value) <= 0 _
           Or CDbl(wsData.Cells(i, "E").Value) <= 0 Then
    
            errorRow = i
            hasError = True
            Exit For
    
        End If

Next i
    
    If hasError Then
        MsgBox errorRow & "行目の入力内容を確認してください。", vbExclamation
        wsData.Activate
        wsData.Cells(errorRow, "A").Select
        Exit Sub
    End If

    For i = 2 To lastRow

        staffName = Trim(wsData.Cells(i, "B").Value)
        productName = Trim(wsData.Cells(i, "C").Value)
        salesAmount = CDbl(wsData.Cells(i, "D").Value) * _
                      CDbl(wsData.Cells(i, "E").Value)
        
        wsData.Cells(i, "F").Value = salesAmount

        If staffName <> "" Then
            If staffDict.Exists(staffName) Then
                staffDict(staffName) = staffDict(staffName) + salesAmount
            Else
                staffDict.Add staffName, salesAmount
            End If
        End If

        If productName <> "" Then
            If productDict.Exists(productName) Then
                productDict(productName) = productDict(productName) + salesAmount
            Else
                productDict.Add productName, salesAmount
            End If
        End If

    Next i

    wsStaff.Range("A4:B" & wsStaff.Rows.Count).ClearContents
    wsProduct.Range("A4:B" & wsProduct.Rows.Count).ClearContents

    outputRow = 4

    For Each key In staffDict.Keys
        wsStaff.Cells(outputRow, "A").Value = key
        wsStaff.Cells(outputRow, "B").Value = staffDict(key)
        outputRow = outputRow + 1
    Next key
    
    staffLastRow = outputRow - 1
    
    wsStaff.Cells(outputRow + 1, "A").Value = "合計"

    wsStaff.Cells(outputRow + 1, "B").Formula = "=SUM(B4:B" & outputRow - 1 & ")"
    
    wsStaff.Cells(outputRow + 1, "A").Font.Bold = True
    wsStaff.Cells(outputRow + 1, "B").Font.Bold = True

    outputRow = 4

    For Each key In productDict.Keys
        wsProduct.Cells(outputRow, "A").Value = key
        wsProduct.Cells(outputRow, "B").Value = productDict(key)
        outputRow = outputRow + 1
    Next key
    
    productLastRow = outputRow - 1

    wsProduct.Cells(outputRow + 1, "A").Value = "合計"
    
    wsProduct.Cells(outputRow + 1, "B").Formula = "=SUM(B4:B" & outputRow - 1 & ")"
    
    wsProduct.Cells(outputRow + 1, "A").Font.Bold = True
    wsProduct.Cells(outputRow + 1, "B").Font.Bold = True

    wsStaff.Columns("A:B").AutoFit
    wsProduct.Columns("A:B").AutoFit
    
    wsStaff.Range("A2").Value = "集計日時：" & Format(Now, "yyyy/mm/dd hh:mm")
    wsProduct.Range("A2").Value = "集計日時：" & Format(Now, "yyyy/mm/dd hh:mm")
    
    wsStaff.Range("A2:B2").Font.Size = 10
    wsProduct.Range("A2:B2").Font.Size = 10
    
        Dim chartObj As ChartObject

    ' 既存グラフを削除
    For Each chartObj In wsStaff.ChartObjects
        chartObj.Delete
    Next chartObj

    For Each chartObj In wsProduct.ChartObjects
        chartObj.Delete
    Next chartObj

    ' 担当者別グラフを作成
    Set chartObj = wsStaff.ChartObjects.Add( _
        Left:=250, _
        Top:=50, _
        Width:=450, _
        Height:=280)

    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsStaff.Range("A3:B" & staffLastRow)
        .HasTitle = True
        .ChartTitle.Text = "担当者別売上"
        .HasLegend = False
    End With

    ' 商品別グラフを作成
    Set chartObj = wsProduct.ChartObjects.Add( _
        Left:=250, _
        Top:=50, _
        Width:=450, _
        Height:=280)

    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsProduct.Range("A3:B" & productLastRow)
        .HasTitle = True
        .ChartTitle.Text = "商品別売上"
        .HasLegend = False
    End With

    wsStaff.Range("B4:B" & wsStaff.Rows.Count).NumberFormatLocal = "#,##0円"
    wsProduct.Range("B4:B" & wsProduct.Rows.Count).NumberFormatLocal = "#,##0円"

    MsgBox "売上集計が完了しました。", vbInformation
    
    Exit Sub
    
ErrorHandler:
        MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
               "エラー番号：" & Err.Number & vbCrLf & _
               "内容：" & Err.Description, vbCritical

End Sub
