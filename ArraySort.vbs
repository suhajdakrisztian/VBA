Public Function SortArray(array_to_sort As Variant, Descending As Boolean)

'Sorts array in specified order 

i = 1
Do Until i = UBound(array_to_sort)
    If Descending = True Then
        If array_to_sort(i + 1) > array_to_sort(i) Then
        pLACEHOLDER = array_to_sort(i)
        array_to_sort(i) = array_to_sort(i + 1)
        array_to_sort(i + 1) = pLACEHOLDER
        i = 0 'should be reset: i+=1 happens at every iteration of the main loop, and if we move elements checking the whole array again is obligatory
        End If
    Else
        If array_to_sort(i) > array_to_sort(i + 1) Then
        pLACEHOLDER = array_to_sort(i)
        array_to_sort(i) = array_to_sort(i + 1)
        array_to_sort(i + 1) = pLACEHOLDER
        i = 0 
        End If
        
    End If
i = i + 1
Loop
SortArray = array_to_sort
End Function


