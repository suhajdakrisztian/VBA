Private Function IsInArray(searched_element As Variant, array As Variant) As Boolean
'Checks if element is in array
On Error GoTo IsInArrayError: 'array is empty
    For Each element In array
        If element = searched_element Then
            IsInArray = True
            Exit Function
        End If
    Next 
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function