Sub ExtractRedParts()
    Dim ws As Worksheet
    Dim r As Long, i As Long
    Dim s As String
    Dim tmp As String
    Dim ch As String
    Dim isRed As Boolean
    Dim parts As Collection
    
    ' 処理高速化用フラグ退避
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    
    ' 高速化モード ON
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set ws = ActiveSheet
    
    ' 列Bの最終行を算出（ソースはB列に貼り付け）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    r = 1
    Do While r <= lastRow
        s = ws.Cells(r, 2).Value   ' B列がソースコード
        Set parts = New Collection
        tmp = ""
        isRed = False
        
        ' 文字ごとに色をチェック
        For i = 1 To Len(s)
            ch = Mid(s, i, 1)
            If ws.Cells(r, 2).Characters(i, 1).Font.Color = vbRed Then
                If Not isRed Then
                    tmp = ch
                    isRed = True
                Else
                    tmp = tmp & ch
                End If
            Else
                If isRed Then
                    parts.Add Trim(tmp)   ' 前後空白を削除
                    tmp = ""
                    isRed = False
                End If
            End If
        Next i
        If isRed Then parts.Add Trim(tmp)   ' 最後が赤字で終わる場合も Trim
        
        ' 赤字部分があれば D列に書き出し
        If parts.Count > 0 Then
            ws.Cells(r, 4).Value = parts(1)
            If parts.Count > 1 Then
                Dim j As Long
                For j = 2 To parts.Count
                    r = r + 1
                    ws.Rows(r).Insert Shift:=xlDown
                    ws.Cells(r, 1).ClearContents   ' A列は空白
                    ws.Cells(r, 4).Value = parts(j)
                    lastRow = lastRow + 1
                Next j
            End If
        End If
        r = r + 1
    Loop
    
    ' 高速化モード OFF
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Sub MarkDuplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim dict As Object
    Dim val As String
    Dim idx As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row  ' D列が抽出対象
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For r = 1 To lastRow
        val = ws.Cells(r, 4).Value ' D列の赤字抽出文字列
        If val <> "" Then
            If Not dict.Exists(val) Then
                idx = dict.Count + 1
                dict.Add val, idx
                ws.Cells(r, 3).Value = "※" & idx   ' C列に付与
            Else
                ws.Cells(r, 3).Value = "※" & dict(val)
            End If
        End If
    Next r
End Sub


Sub ResetExtractedData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    Set ws = ActiveSheet
    
    ' 1. C列とD列をクリア
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range("C1:D" & lastRow).ClearContents
    
    ' 2. A列が空白の行を下から削除
    For r = lastRow To 1 Step -1
        If ws.Cells(r, 1).Value = "" Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub


