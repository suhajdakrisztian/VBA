Function FlexibleSort(array_to_sort, IsDescending As Boolean)
'Created by Krisztián Suhajda
'The algo is based on Bubble sort
'the result array must be instantiated as a new variable
'Arrays may either start at 0 or 1
    array_length = UBound(array_to_sort)
    If LBound(array_to_sort) = 0 Then
    For i = 0 To array_length
        For j = 0 To array_length - i - 1:
            If IsDescending Then
                If array_to_sort(j) < array_to_sort(j + 1) Then
                    temp = array_to_sort(j)
                    array_to_sort(j) = array_to_sort(j + 1)
                    array_to_sort(j + 1) = temp
                End If
            Else
                If array_to_sort(j) > array_to_sort(j + 1) Then
                        temp = array_to_sort(j)
                        array_to_sort(j) = array_to_sort(j + 1)
                        array_to_sort(j + 1) = temp
                End If
            End If
        Next
    Next
    FlexibleSort = array_to_sort
    ElseIf LBound(array_to_sort) <> 0 Then
        For i = LBound(array_to_sort) To array_length + 1
            For j = LBound(array_to_sort) To array_length - 1
                If IsDescending Then
                    If array_to_sort(j) < array_to_sort(j + 1) Then
                        temp = array_to_sort(j)
                        array_to_sort(j) = array_to_sort(j + 1)
                        array_to_sort(j + 1) = temp
                    End If
                Else
                    If array_to_sort(j) > array_to_sort(j + 1) Then
                            temp = array_to_sort(j)
                            array_to_sort(j) = array_to_sort(j + 1)
                            array_to_sort(j + 1) = temp
                    End If
                End If
            Next
        Next
    Else
    FlexibleSort = "ENTER A VALID START" 'Hope this will never be visible
    End If
    FlexibleSort = array_to_sort
End Function