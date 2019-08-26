Sub Fill_Range_With_Formula():

    Application.ScreenUpdating = False

    ActiveCell.Copy
    iterable_row = ActiveCell.Row
    Base_column = ActiveCell.Column
    
    Do While Cells(iterable_row + 1, ActiveCell.Column - 1) <> ""
        Cells(iterable_row + 1, Base_column).PasteSpecial xlFormulas
        iterable_row = iterable_row + 1
    Loop
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
End Sub

