Sub Fill_Range_With_Formula():
    Application.ScreenUpdating = False
    static_row = ActiveCell.Row
    first_column = ActiveCell.Column
    last_column = ActiveCell.End(xlToRight).Column
    For column_number = first_column To last_column
        Cells(static_row, column_number).Activate
        ActiveCell.Copy
        iterable_row = ActiveCell.Row
        Base_column = ActiveCell.Column
        Do While Cells(iterable_row + 1, ActiveCell.Column - 1) <> ""
            Cells(iterable_row + 1, Base_column).PasteSpecial xlFormulas
            iterable_row = iterable_row + 1
        Loop   
    Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

