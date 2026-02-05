Kodik

    Sub TransferFromJtoNOP()
    
        Dim ws As Worksheet
        Set ws = ActiveSheet
    
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
        Dim i As Long
        Dim cyclePos As Integer
        Dim baseRow As Long
    
        cyclePos = 1
    
        For i = 1 To lastRow
    
            If ws.Cells(i, "J").Value <> "" Then
    
                ' начало нового цикла
                If cyclePos = 1 Then
                    baseRow = i
                    ws.Cells(baseRow, "N").Value = ws.Cells(i, "J").Value
                ElseIf cyclePos = 2 Then
                    ws.Cells(baseRow, "O").Value = ws.Cells(i, "J").Value
                ElseIf cyclePos = 3 Then
                    ws.Cells(baseRow, "P").Value = ws.Cells(i, "J").Value
                End If
    
                cyclePos = cyclePos + 1
    
                ' сброс после 3 строк
                If cyclePos > 3 Then
                    cyclePos = 1
                End If
    
                ' сброс по тонкой нижней границе
                If ws.Cells(i, "J").Borders(xlEdgeBottom).LineStyle = xlContinuous _
                   And ws.Cells(i, "J").Borders(xlEdgeBottom).Weight = xlThin Then
                    cyclePos = 1
                End If
    
            End If
        Next i
    
        MsgBox "Готово 👍"
    
    End Sub



konec kodica
