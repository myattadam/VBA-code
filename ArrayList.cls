VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Enum SORTBY
    DESCENDING
    ASCENDING
End Enum

Dim data As Variant

Property Get List() As Variant
    List = data
End Property

Property Get Length() As Long
    Length = UBound(data) - LBound(data) + 1
End Property

Property Get value(index As Variant) As Variant
    Attribute Value.VB_UserMemId = 0
    value = data(index)
End Property

Property Let value(index As Variant, value_ As Variant)
    Attribute Value.VB_UserMemId = 0
    If index > Length Then ReDim Preserve data(1 To index)
    data(index) = value_
End Property


Sub AddTo(arr As Variant, value As Variant)
    If IsEmpty(arr) Then
        ReDim arr(1 To 1)
    Else
        ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
    End If
    
    arr(UBound(arr)) = value
End Sub

Function Add(value As Variant) As Long
    AddTo data, value
    Add = UBound(data)
End Function

Sub Clear()
    If IsArray(data) Then Erase data Else data = Empty
End Sub


Private Function Compare(method As SORTBY, lhs As Variant, rhs As Variant) As Boolean
    Select Case method
        Case SORTBY.DESCENDING: Compare = lhs < rhs
        Case SORTBY.ASCENDING: Compare = lhs > rhs
    End Select
End Function


Function QuickSort(arr As Variant, method As SORTBY, Optional byColumn As Long = -1) As Variant
    Dim left As Variant
    Dim right As Variant
    Dim pivot As Variant
    
    Dim i As Long
    
    If Arrays.Length(arr) > 1 Then
    
        pivot = arr(UBound(arr))
        
        For i = LBound(arr) To UBound(arr) - 1
            If byColumn = -1 Then
        
                If Compare(method, pivot, arr(i)) Then
                    Arrays.Add left, arr(i)
                Else
                    Arrays.Add right, arr(i)
                End If
                
            Else
            
                If Compare(method, pivot(byColumn), arr(i)(byColumn)) Then
                    Arrays.Add left, arr(i)
                Else
                    Arrays.Add right, arr(i)
                End If
                
            End If
        Next
        
        QuickSort left, method, byColumn
        QuickSort right, method, byColumn
    
        arr = Empty
        
        If Not IsEmpty(left) Then
            For i = LBound(left) To UBound(left)
                Arrays.Add arr, left(i)
            Next
        End If
        
        Arrays.Add arr, pivot
        
        If Not IsEmpty(right) Then
            For i = LBound(right) To UBound(right)
                Arrays.Add arr, right(i)
            Next
        End If
        
    End If
    
    QuickSort = arr
End Function












'Property Get Sum() As Variant
'    Dim i As Long
'
'    Dim value As Variant
'    value = 0
'
'    For i = LBound(data) To UBound(data)
'
'        If IsArray(data(i)) Then
'
'            value = value + Sum(data(i))
'
'        Else
'
'            Select Case TypeName(data(i))
'                Case "Integer", "Long", "Single", "Double"
'                    value = value + data(i)
'            End Select
'
'        End If
'
'    Next
'
'    Sum = value
'
'End Property
'
'
'
'Property Get Average() As Double
'    Dim i As Long
'
'    Dim value As Variant
'    Dim Count As Long
'    value = 0
'
'    For i = LBound(data) To UBound(data)
'
'        Select Case TypeName(data(i))
'            Case "Integer", "Long", "Single", "Double"
'                value = value + data(i)
'                Count = Count + 1
'        End Select
'
'    Next
'
'    Average = value / Count
'
'End Property



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
'    arr = QuickSort(arr, 2)
'    arr = QuickSort(arr, 0)
'
'    Debug.Print ""
'
'    For i = LBound(arr) To UBound(arr)
'        Debug.Print arr(i)(0); vbTab; arr(i)(1); vbTab; CDate(arr(i)(2))
'    Next
'
'    Dim fub As Variant
'    ReDim fub(1 To 5)
'
'    Clear fub
'
'    Debug.Print "Done"
'
'End Sub
'
'

