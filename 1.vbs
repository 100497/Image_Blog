Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Static OldRange As Range
    Static OldColor As Variant
    
    ' 恢复之前的颜色
    If Not OldRange Is Nothing Then
        OldRange.Interior.Color = OldColor
    End If
    
    ' 保存当前选择的单元格范围和颜色
    Set OldRange = Target
    OldColor = Target.Interior.Color
    
    ' 突出显示当前行和列
    With Target
        .EntireRow.Interior.Color = RGB(255, 255, 0) ' 黄色
        .EntireColumn.Interior.Color = RGB(255, 255, 0) ' 黄色
    End With
End Sub
