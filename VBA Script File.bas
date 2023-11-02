Attribute VB_Name = "Module1"
Sub main()
        Dim sh As Worksheet
        For Each sh In ThisWorkbook.Sheets
                If sh.Name <> "program" Then
                        Call get_results_for_sheets(sh)
                        sh.Range("K:K, P2:P3").NumberFormat = "0.00%"
                        Call ApplyConditionalFormatting(sh)
                End If
        Next sh
End Sub
Sub get_results_for_sheets(sh As Worksheet)
        'read array
        Dim arr As Variant: arr = sh.Range("A1").CurrentRegion.Value
        Call cleanArr(arr)
        Dim yr As Integer: yr = Year(arr(2, 2))
        
        'get unique tickers
        Dim tickersArr As Variant: tickersArr = get_all_tickers(arr)
        
        ReDim Preserve tickersArr(1 To UBound(tickersArr), 1 To 4)
        Dim i As Long
        For i = 1 To UBound(tickersArr, 1)
                Dim ticker As String: ticker = tickersArr(i, 1)
                Dim changeArr As Variant: changeArr = price_change_vol(arr, ticker, yr)
                tickersArr(i, 2) = changeArr(1, 1)
                tickersArr(i, 3) = changeArr(2, 1)
                tickersArr(i, 4) = changeArr(3, 1)
        Next i
        
        Call write_data(tickersArr, sh)
        Call get_greatest_info(tickersArr, sh)
End Sub

Function get_all_tickers(arr) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To UBound(arr, 1)
        If Not dict.Exists(arr(i, 1)) Then
            dict.Add arr(i, 1), 1
        End If
    Next i
    
    Dim uniqueArr() As Variant
    ReDim uniqueArr(1 To dict.Count, 1 To 1)
    
    
    Dim rowIndex As Long: rowIndex = 1
    Dim key As Variant
    For Each key In dict.Keys
        uniqueArr(rowIndex, 1) = key
        rowIndex = rowIndex + 1
    Next key
    
    get_all_tickers = uniqueArr
End Function

Sub cleanArr(arr As Variant)

        Dim i As Long
        For i = 2 To UBound(arr, 1)
                Dim yr As Integer: yr = CInt(Left(arr(i, 2), 4))
                Dim mnth As Integer: mnth = CInt(Mid(arr(i, 2), 5, 2))
                Dim dy As Integer: dy = CInt(Right(arr(i, 2), 2))
                arr(i, 2) = DateSerial(yr, mnth, dy)
        Next i
End Sub

Function price_change_vol(arr As Variant, ticker As String, yr As Integer) As Variant
        Dim temp(1 To 3, 1 To 1) As Variant
        
        Dim openPrice As Double, closePrice As Double
        Dim i As Long
        Dim foundOpen As Boolean: foundOpen = False
        Dim vol As LongLong: vol = 0
        For i = 2 To UBound(arr, 1)
                If Year(arr(i, 2)) > yr Then Exit For
                If Year(arr(i, 2)) = yr And arr(i, 1) = ticker Then
                        If foundOpen = False Then
                                openPrice = arr(i, 3)
                                foundOpen = True
                        End If
                        
                        closePrice = arr(i, 6)
                        
                        vol = vol + arr(i, 7)
                End If
        Next i
        
        temp(1, 1) = closePrice - openPrice
        temp(2, 1) = (closePrice - openPrice) / openPrice
        temp(3, 1) = vol
        price_change_vol = temp
End Function

Sub get_greatest_info(arr As Variant, sh As Worksheet)
        Dim result(1 To 4, 1 To 3) As Variant
        Dim gIncTicker As String, gDecTicker As String, gVolTicker As String
        Dim gInc As Double: gInc = 0
        Dim gDec As Double: gDec = 0
        Dim gVol As LongLong: gVol = 0
        Dim i As Long
        For i = 1 To UBound(arr, 1)
                If arr(i, 3) < gDec Then
                        gDec = arr(i, 3)
                        gDecTicker = arr(i, 1)
                End If
                If arr(i, 3) > gInc Then
                        gInc = arr(i, 3)
                        gIncTicker = arr(i, 1)
                End If
                If arr(i, 4) > gVol Then
                        gVol = arr(i, 4)
                        gVolTicker = arr(i, 1)
                End If
        Next i
        
        result(1, 2) = "TICKERS"
        result(1, 3) = "VALUES"
        result(2, 1) = "Greatest increase"
        result(2, 2) = gIncTicker
        result(2, 3) = gInc
        result(3, 1) = "Greatest decrease"
        result(3, 2) = gDecTicker
        result(3, 3) = gDec
        result(4, 1) = "Greatest Vol"
        result(4, 2) = gVolTicker
        result(4, 3) = gVol
        
        With sh
                .Range("O1").CurrentRegion.ClearContents
                .Range("N1").Resize(UBound(result, 1), UBound(result, 2)) = result
        End With
        
End Sub

Sub write_data(result As Variant, sh As Worksheet)
        With sh
                .Range("i1").CurrentRegion.ClearContents
                .Range("i1").Value = "TICKERS"
                .Range("j1").Value = "YEARLY CHANGE"
                .Range("k1").Value = "PERCENT CHANGE"
                .Range("l1").Value = "TOTAL VOLUME"
                .Range("i2").Resize(UBound(result, 1), UBound(result, 2)) = result
        End With
End Sub

Sub ApplyConditionalFormatting(sh As Worksheet)
    
    Dim rng As Range
    Set rng = sh.Range("$I:$L")
    
    ' Clear existing conditional formatting in the range
    rng.FormatConditions.Delete
    
    ' Define the green formatting rule
    Dim greenRule As FormatCondition
    Set greenRule = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($J1>0, ISNUMBER($J1))")
    greenRule.Interior.Color = RGB(0, 255, 0) ' Green color
    
    ' Define the red formatting rule
    Dim redRule As FormatCondition
    Set redRule = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($J1<0, ISNUMBER($J1))")
    redRule.Interior.Color = RGB(255, 0, 0) ' Red color
    
End Sub

