Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim rngTable As Range
    Dim rngHeaders As Range
    Dim rngLabels As Range
    Dim rngData As Range
    Dim rngTotalRow As Range
    Dim cell As Range
    Dim i As Long
    
    Application.ScreenUpdating = False
    Set ws = Target.Parent
    
    ' 전체 시트 초기화
    With ws.Cells
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone
        .Font.Name = "Segoe UI"
        .Font.Size = 10
    End With
    
    For Each pt In ws.PivotTables
        Set rngTable = pt.TableRange2
        '▼▼▼ 변경된 부분 ▼▼▼
        Set rngHeaders = pt.TableRange2.Rows(1) ' 첫 번째 행을 헤더로 강제 지정
        '▲▲▲ 변경된 부분 ▲▲▲
        Set rngLabels = pt.RowRange
        Set rngData = pt.DataBodyRange
        
        ' 헤더 영역 스타일링 (첫 번째 행 강제 적용)
        If Not rngHeaders Is Nothing Then
            With rngHeaders
                .Interior.Color = RGB(204, 229, 255)
                .Font.Bold = True
                .Font.Color = RGB(0, 51, 102)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(153, 153, 153)
                .RowHeight = 28
            End With
        End If
        
        ' 행 레이블 스타일링
        If Not rngLabels Is Nothing Then
            With rngLabels
                .Interior.Color = RGB(245, 245, 245)
                .Font.Bold = True
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Color = RGB(221, 221, 221)
            End With
        End If
        
        ' 데이터 영역 스타일링
        If Not rngData Is Nothing Then
            With rngData
                .Interior.Color = RGB(255, 255, 255)
                .HorizontalAlignment = xlRight
                .NumberFormat = "#,##0"
                
                ' 대체 행 색상 (줄무늬 효과)
                For i = 1 To .Rows.Count Step 2
                    .Rows(i).Interior.Color = RGB(249, 249, 249)
                Next i
                
                ' 테두리 설정
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(221, 221, 221)
            End With
        End If
        
        ' 합계 행 스타일링 (모든 총계 행 탐색)
        On Error Resume Next
        '▼▼▼ 변경된 부분 ▼▼▼
        Set rngTotalRow = pt.TableRange2.Rows(pt.TableRange2.Rows.Count)
        If InStr(1, rngTotalRow.Cells(1, 1).Text, "Total", vbTextCompare) > 0 Then
            With rngTotalRow
                .Interior.Color = RGB(204, 229, 255)
                .Font.Bold = True
                .Borders(xlTop).LineStyle = xlDouble
                .Borders(xlTop).Color = RGB(0, 51, 102)
            End With
        End If
        '▲▲▲ 변경된 부분 ▲▲▲
        On Error GoTo 0
        
        
        ' 전체 테이블 외곽 테두리
        With rngTable.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = RGB(0, 51, 102)
            .Weight = xlThick
        End With
        With rngTable.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(0, 51, 102)
            .Weight = xlThick
        End With
    Next pt
    
    ' 열 너비 및 행 높이 최적화
    With ws
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        For Each cell In rngTable.Columns
            cell.ColumnWidth = cell.ColumnWidth * 1.1 ' 10% 여유 확보
        Next cell
    End With
    
    Application.ScreenUpdating = True
End Sub

