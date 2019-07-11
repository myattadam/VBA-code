Attribute VB_Name = "Arrays"
'Enum SORTBY
'    DESCENDING
'    ASCENDING
'End Enum
'
'Public Function IsIn(ByVal source As Variant, match As Variant) As Boolean
'    If IsArray(match) Then Err.Raise 1000, , "Cannot search using an array."
'
'    If IsArray(source) Then
'        Dim comp As Variant
'
'        For Each comp In source
'            If Not IsArray(comp) Then
'                If comp = match Then
'                    IsIn = True
'                    Exit Function
'                End If
'            End If
'        Next
'    End If
'
'    IsIn = False
'End Function
'
'
''Function Subtract(A As Variant, B As Variant) As Variant
''    Dim result As Variant
''
''    ReDim result(WorksheetFunction.max(LBound(A), LBound(B)) To WorksheetFunction.min(UBound(A), UBound(B)))
''
''    For i& = LBound(result) To UBound(result)
''        result(i) = A(i) - B(i)
''    Next
''
''    ArraySubtract = result
''End Function
''
''
''Function Add(A As Variant, B As Variant) As Variant
''    Dim result As Variant
''
''    ReDim result(WorksheetFunction.min(LBound(A), LBound(B)) To WorksheetFunction.max(UBound(A), UBound(B)))
''
''    For i& = LBound(result) To UBound(result)
''        On Error Resume Next
''        result(i) = 0
''        result(i) = result(i) + A(i)
''        result(i) = result(i) + B(i)
''    Next
''
''    ArrayAdd = result
''End Function
'
'
'Function Add(ByRef arr As Variant, value As Variant)
'    If IsEmpty(arr) Then
'        ReDim arr(1 To 1)
'    Else
'        ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
'    End If
'
'    arr(UBound(arr)) = value
'End Function
'
'
'Function Length(ByRef arr As Variant) As Long
'    Length = 0
'
'    If IsArray(arr) Then Length = UBound(arr) - LBound(arr) + 1
'End Function
'
'
'Function Clear(ByRef arr As Variant)
'    If IsArray(arr) Then Erase arr Else arr = Empty
'End Function
'
''Function Compare(ByVal arr As Variant, operator As String) As Long
''    Dim i As Long
''
''    Compare = LBound(arr)
''
''    For i = LBound(arr) To UBound(arr)
''        If Evaluate(arr(i) & operator & arr(Compare)) Then Compare = i
''    Next
''End Function
'
''''' REWRITE FOR STAGGERED ARRAYS - NOT NEEDED - CHECK FOR DEPENDANCIES
''Function Smallest(ByVal arr As Variant) As Long
''    Dim i As Long
''
''    Smallest = LBound(arr)
''
''    For i = LBound(arr) To UBound(arr)
''        If arr(i) < arr(Smallest) Then Smallest = i
''    Next
''End Function
''
''''' REWRITE FOR STAGGERED ARRAYS - NOT NEEDED - CHECK FOR DEPENDANCIES
''Function Largest(ByVal arr As Variant) As Long
''    Dim i As Long
''
''    Largest = LBound(arr)
''
''    For i = LBound(arr) To UBound(arr)
''        If arr(i) > arr(Largest) Then Largest = i
''    Next
''End Function
'
''''' REWRITE FOR STAGGERED ARRAYS - NOT NEEDED - CHECK FOR DEPENDANCIES
''Function Average(ByVal arr As Variant) As Variant
''    Dim i As Long
''    Dim sum As Variant
''
''    For i = LBound(arr) To UBound(arr)
''        sum = sum + arr(i)
''    Next
''
''    Average = sum / Length(arr)
''End Function
'
'
'
'Function Sum(arr As Variant, Optional byColumn As Long = -1) As Variant
'    Dim i As Long
'
'    For i = LBound(arr) To UBound(arr)
'        If byColumn > -1 Then
'            Sum = Sum + arr(i)(byColumn)
'        Else
'            Sum = Sum + arr(i)
'        End If
'    Next
'
'End Function
'
'' Returns an array length 2, containing the slope and Y-intercept
'Function Trend(ByVal pair As Variant) As Variant
'    Const X As Long = 1
'    Const Y As Long = 2
'
'    Dim i As Long
'    Dim lower As Long
'    Dim upper As Long
'    Dim size As Long
'    Dim mean(1 To 2) As Double
'
'    lower = LBound(pair)
'    upper = UBound(pair)
'    size = upper - lower + 1
'
'    For i = lower To upper
'        mean(X) = mean(X) + pair(i)(X)
'        mean(Y) = mean(Y) + pair(i)(Y)
'    Next
'
'    mean(X) = mean(X) / size
'    mean(Y) = mean(Y) / size
'
'    Dim temp As Variant
'    ReDim temp(LBound(pair) To UBound(pair))
'
'    For i = lower To upper
'        temp(i) = Array((pair(i)(X) - mean(X)) * (pair(i)(Y) - mean(Y)), _
'                        (pair(i)(X) - mean(X)) ^ 2)
'    Next
'
'    Dim Sum(1 To 2) As Variant
'
'    For i = lower To upper
'        Sum(1) = Sum(1) + temp(i)(0)
'        Sum(2) = Sum(2) + temp(i)(1)
'    Next
'
'    Sum(1) = Sum(1) / size
'    Sum(2) = Sum(2) / size
'
'    Dim result(1 To 2) As Double
'        result(1) = Sum(1) / Sum(2) ' Slope
'        result(2) = mean(Y) - (mean(X) * result(1)) ' Y-intercept
'
'    Trend = result
'End Function
'
'''' arr = either a one-dimensional array or a staggered array
'''' byColumn = for staggered arrays, specify the column to sort by
'
'Function Compare(method As SORTBY, lhs As Variant, rhs As Variant) As Boolean
'    Select Case method
'        Case SORTBY.DESCENDING: Compare = lhs < rhs
'        Case SORTBY.ASCENDING: Compare = lhs > rhs
'    End Select
'End Function
'
'Function QuickSort(arr As Variant, method As SORTBY, Optional byColumn As Long = -1) As Variant
'    Dim left As Variant
'    Dim right As Variant
'    Dim pivot As Variant
'
'    Dim i As Long
'
'    If Arrays.Length(arr) > 1 Then
'
'        pivot = arr(UBound(arr))
'
'        For i = LBound(arr) To UBound(arr) - 1
'            If byColumn = -1 Then
'
'                If Compare(method, pivot, arr(i)) Then
'                    Arrays.Add left, arr(i)
'                Else
'                    Arrays.Add right, arr(i)
'                End If
'
'            Else
'
'                If Compare(method, pivot(byColumn), arr(i)(byColumn)) Then
'                    Arrays.Add left, arr(i)
'                Else
'                    Arrays.Add right, arr(i)
'                End If
'
'            End If
'        Next
'
'        QuickSort left, method, byColumn
'        QuickSort right, method, byColumn
'
'        arr = Empty
'
'        If Not IsEmpty(left) Then
'            For i = LBound(left) To UBound(left)
'                Arrays.Add arr, left(i)
'            Next
'        End If
'
'        Arrays.Add arr, pivot
'
'        If Not IsEmpty(right) Then
'            For i = LBound(right) To UBound(right)
'                Arrays.Add arr, right(i)
'            Next
'        End If
'
'    End If
'
'    QuickSort = arr
'End Function
'
'
'Sub test()
'
'    Dim arr As Variant
'
'    For i = 0 To 15
'        num = Int(Rnd() * 5)
'        code = ChrW(65 + Int(Rnd() * 26))
'
'        value = Round(Rnd() * 100, 2)
'        d = #1/1/2019# + Rnd() * 365
'
'        Arrays.Add arr, Array(code, value, CDbl(d))
'    Next
'
'    Arrays.Add arr, Array("#", 0, 0)
'
'
'    Debug.Print ""
'
'    For i = LBound(arr) To UBound(arr)
'        Debug.Print arr(i)(0); vbTab; arr(i)(1); vbTab; CDate(arr(i)(2))
'    Next
'
'    arr = QuickSort(arr, DESCENDING, 0)
'
'    Debug.Print ""
'
'    For i = LBound(arr) To UBound(arr)
'        Debug.Print arr(i)(0); vbTab; arr(i)(1); vbTab; CDate(arr(i)(2))
'    Next
'
'    Debug.Print "Done"
'
'End Sub
