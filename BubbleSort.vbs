Function BubbleSort(array_to_sort, IsDescending As Boolean)

'Created b Krisztián Suhajda
'Prerequisites:
'Array contains numeric values
'Array starts at 0
'the new array must be instantiated as a new variable


    array_length = UBound(array_to_sort)

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
    
    BubbleSort = array_to_sort
    
End Function
