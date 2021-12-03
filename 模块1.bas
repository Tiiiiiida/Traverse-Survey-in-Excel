Attribute VB_Name = "Module1"
Sub InputNum_Click()
    '清空当前Sheet
    Cells.Clear

    '开始生成表头
    Range("A1:A3").Merge
    Range("A1").Value = "点号"

    Range("B1:B2").Merge
    Range("B1").Value = "水平角度"
    Range("B3").Value = "°  ′  ″"

    Range("C1:C2").Merge
    Range("C1").Value = "改正数"
    Range("C3").Value = "″"

    Range("D1:D2").Merge
    Range("D1").Value = "改正后水平角度"
    Range("D3").Value = "°  ′  ″"

    Range("E1:E2").Merge
    Range("E1").Value = "坐标方位角"
    Range("E3").Value = "°  ′  ″"

    Range("F1:F2").Merge
    Range("F1").Value = "距离"
    Range("F3").Value = "m"

    Range("G1:H1").Merge
    Range("G2:H2").Merge
    Range("G1").Value = "坐标增量"
    Range("G2").Value = "m"
    Range("G3").Value = "△x"
    Range("H3").Value = "△y"

    Range("I1:J1").Merge
    Range("I2:J2").Merge
    Range("I1").Value = "改正后坐标增量"
    Range("I2").Value = "m"
    Range("I3").Value = "△x"
    Range("J3").Value = "△y"

    Range("K1:L1").Merge
    Range("K2:L2").Merge
    Range("K1").Value = "坐标"
    Range("K2").Value = "m"
    Range("K3").Value = "x"
    Range("L3").Value = "y"

    Range("A4:A5").Merge
    Range("A4").Value = "总计"
    Range("A6:B6").Merge
    Range("A6").Value = "角 度 闭 合 差："
    Range("A8:B8").Merge
    Range("A8").Value = "坐标增量闭合差："
    Range("A10:B10").Merge
    Range("A10").Value = "导线全长相对闭合差："
    '填写内容完成

    With Range("A1:L10")        '单元格对齐方式改为居中
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With Range("A1:L5")      '为上部分添加所有框线
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With

    With Range("A6:L10")        '为下部分添加外框线
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With

    With Range("B:B").ColumnWidth = 14      '调整列宽
    End With
    With Range("D:D").ColumnWidth = 14
    End With
    With Range("E:E").ColumnWidth = 14
    End With
    With Range("F:L").ColumnWidth = 9.7
    End With
    '表头生成完成

    '开始依据输入的测量点数量生成表格内部行
    Dim RowNum As Integer       '输入RowNum
        RowNum = InputBox("Please input the num of points.", "Num of Points")

    Dim i As Integer
    For i = 1 To 2 * RowNum + 2      '添加所需要的行，每一个测量点需要两行
        Rows("4:4").Insert
    Next

    For i = 4 To 2 * RowNum + 6 Step 2      '合并单元格 左半部分
        For j = 1 To 4
        Range(Cells(i, j), Cells(i + 1, j)).Merge
        Next
    Next
    For i = 5 To 2 * RowNum + 5 Step 2      '合并单元格 右半部分
        For j = 5 To 12
        Range(Cells(i, j), Cells(i + 1, j)).Merge
        Next
    Next
    '行生成完成

End Sub
