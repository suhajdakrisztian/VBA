Function LastRow(Optional TargetColumn As Long = ActiveCell.Column) As Long
    LastRow = ActiveSheet.Cells(Rows.Count, TargetColumn).End(xlUp).Row
End Function

Function LastColumn(Optional TargetRow As Long = ActiveCell.Row) As Long
    LastColumn = ActiveSheet.Cells(TargetRow, Columns.Count).End(xlToLeft).Column
End Function

