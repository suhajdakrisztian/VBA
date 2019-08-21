Function BubbleSort(array_to_sort)

    array_length = UBound(array_to_sort)
    For i = 0 To array_length
        For j = 0 To array_length - i - 1:
            
            If array_to_sort(j) > array_to_sort(j + 1) Then
                    temp = array_to_sort(j)
                    array_to_sort(j) = array_to_sort(j + 1)
                    array_to_sort(j + 1) = temp
            End If
        Next
    Next
    
    BubbleSort = array_to_sort
    
End Function
