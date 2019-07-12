Attribute VB_Name = "Offcuts"
'@Folder("VBAProject")

Public Enum SortCollectionBy
    DESCENDING
    ASCENDING
End Enum

Sub BubbleSortCollection(c As Collection)

    Do
        swapped = False

        For i = 2 To c.count
            If c(i - 1) > c(i) Then
                temp = c(i)

                c.Remove i
                c.Add temp, , i - 1

                swapped = True
            End If
    Next

    Loop While swapped

End Sub


Private Function Compare(method As SortCollectionBy, lhs As Variant, rhs As Variant) As Boolean
    On Error Resume Next
    Select Case method
        Case SortCollectionBy.DESCENDING: Compare = lhs < rhs
        Case SortCollectionBy.ASCENDING: Compare = lhs > rhs
    End Select
End Function

Function SortCollection(c As Collection, method As SortCollectionBy, Optional byColumn As Long = -1) As Collection
    Dim left As New Collection
    Dim right As New Collection
    Dim pivot As Variant, item As Variant
    
    Dim i As Long
    
    If c.count > 1 Then
    
        pivot = c(c.count)
        
        For i = 1 To c.count - 1
            If byColumn = -1 Then
        
                If Compare(method, pivot, c(i)) Then
                    left.Add c(i)
                Else
                    right.Add c(i)
                End If
                
            Else
            
                If Compare(method, pivot(byColumn), c(i)(byColumn)) Then
                    left.Add c(i)
                Else
                    right.Add c(i)
                End If
                
            End If
        Next
        
        SortCollection left, method, byColumn
        SortCollection right, method, byColumn
    
        Set c = New Collection
        
        For Each item In left
            c.Add item
        Next
        
        c.Add pivot
        
        For Each item In right
            c.Add item
        Next
        
    End If
    
    Set SortCollection = c
End Function
