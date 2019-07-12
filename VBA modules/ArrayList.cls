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

'Dim data As Variant
Public data As Variant

' --- Enums ----------------------------------------------------------------------------------------------

Public Enum SortArrayBy
    DESCENDING
    ASCENDING
End Enum

Public Enum CALC
    sum
    Average
    Multiply
    Divide
End Enum

' --- Default properties ---------------------------------------------------------------------------------

Property Get value(index As Variant) As Variant
Attribute value.VB_UserMemId = 0
    value = data(index)
End Property

Property Let value(index As Variant, value_ As Variant)
Attribute value.VB_UserMemId = 0
    If index > Length Then ReDim Preserve data(1 To index)
    
    If IsObject(value_) Then
        Set data(index) = value_
    Else
        data(index) = value_
    End If
End Property



' --- Class properties -----------------------------------------------------------------------------------

Property Get List() As Variant
    List = data
End Property

Property Get Length() As Long
    Length = LengthOf(data)
End Property



' --- Class functions ------------------------------------------------------------------------------------

Function Add(value As Variant, Optional unzipArray As Boolean = False) As Long
    If IsArray(value) And unzipArray Then
        Dim i As Long
        
        For i = LBound(value) To UBound(value)
            AddTo data, value(i)
        Next
    Else
        AddTo data, value
    End If
    
    Add = UBound(data)
End Function

Function Calculate(method As CALC, Optional value As Variant = 1) As Variant
    Select Case method
        Case CALC.sum: Calculate = SumOf(data)
        Case CALC.Average: Calculate = AverageOf(data)
        Case CALC.Multiply: MultiplyBy data, value
        Case CALC.Divide: DivideBy data, value
    End Select
End Function

Sub ApplyFormat(fmt As String)
    FormatAs data, fmt
End Sub

Sub Sort(method As SortArrayBy, Optional byColumn As Long = -1)
    QuickSort data, method, byColumn
End Sub


Sub Clear()
    'If IsArray(data) Then Erase data Else data = Empty
    data = Empty
End Sub

Sub Transpose()
    data = WorksheetFunction.Transpose(data)
End Sub




Private Sub AddTo(arr As Variant, value As Variant)
    If IsEmpty(arr) Then
        ReDim arr(1 To 1)
    Else
        ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
    End If
    
    arr(UBound(arr)) = value
End Sub



Private Function LengthOf(arr As Variant) As Long
    LengthOf = 0
    If IsArray(arr) Then LengthOf = UBound(arr) - LBound(arr) + 1
End Function

Private Sub FormatAs(arr As Variant, fmt As String)
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)

        If IsArray(arr(i)) Then
            FormatAs arr(i), fmt
        Else
            arr(i) = Format(arr(i), fmt)
        End If

    Next
End Sub

Private Function SumOf(arr As Variant, Optional ByRef value As Variant = 0) As Variant
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)

        If IsArray(arr(i)) Then
            SumOf arr(i), value
        Else
            Select Case TypeName(arr(i))
                Case "Integer", "Long", "Single", "Double"
                    value = value + arr(i)
            End Select
        End If

    Next

    SumOf = value
End Function

Private Function AverageOf(arr As Variant, Optional ByRef value As Variant = 0, Optional ByRef count As Long = 0) As Variant
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)

        If IsArray(arr(i)) Then
            AverageOf arr(i), value, count
        Else
            Select Case TypeName(arr(i))
                Case "Integer", "Long", "Single", "Double"
                    value = value + arr(i)
                    count = count + 1
            End Select
        End If

    Next

    AverageOf = value / count
End Function

Private Sub MultiplyBy(arr As Variant, value As Variant)
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If IsArray(arr(i)) Then
            MultiplyBy arr(i), value
        Else
            Select Case TypeName(arr(i))
                Case "Integer", "Long", "Single", "Double"
                    arr(i) = arr(i) * value
            End Select
        End If
    Next
End Sub

Private Sub DivideBy(arr As Variant, value As Variant)
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If IsArray(arr(i)) Then
            MultiplyBy arr(i), value
        Else
            Select Case TypeName(arr(i))
                Case "Integer", "Long", "Single", "Double"
                    arr(i) = arr(i) / value
            End Select
        End If
    Next
End Sub

Private Function PercentOf(arr As Variant, Optional ByRef value As Variant = 0, Optional ByRef count As Long = 0) As Variant
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)

        If IsArray(arr(i)) Then
            AverageOf arr(i), value, count
        Else
            Select Case TypeName(arr(i))
                Case "Integer", "Long", "Single", "Double"
                    value = value + arr(i)
                    count = count + 1
            End Select
        End If

    Next

    AverageOf = value / count
End Function



Private Function Compare(method As SortArrayBy, lhs As Variant, rhs As Variant) As Boolean
    On Error Resume Next
    Select Case method
        Case SortArrayBy.DESCENDING: Compare = lhs < rhs
        Case SortArrayBy.ASCENDING: Compare = lhs > rhs
    End Select
End Function

Private Function QuickSort(arr As Variant, method As SortArrayBy, Optional byColumn As Long = -1) As Variant
    Dim left As Variant
    Dim right As Variant
    Dim pivot As Variant
    
    Dim i As Long
    
    If LengthOf(arr) > 1 Then
    
        pivot = arr(UBound(arr))
        
        For i = LBound(arr) To UBound(arr) - 1
            If byColumn = -1 Then
        
                If Compare(method, pivot, arr(i)) Then
                    AddTo left, arr(i)
                Else
                    AddTo right, arr(i)
                End If
                
            Else
            
                If Compare(method, pivot(byColumn), arr(i)(byColumn)) Then
                    AddTo left, arr(i)
                Else
                    AddTo right, arr(i)
                End If
                
            End If
        Next
        
        QuickSort left, method, byColumn
        QuickSort right, method, byColumn
    
        arr = Empty
        
        If Not IsEmpty(left) Then
            For i = LBound(left) To UBound(left)
                AddTo arr, left(i)
            Next
        End If
        
        AddTo arr, pivot
        
        If Not IsEmpty(right) Then
            For i = LBound(right) To UBound(right)
                AddTo arr, right(i)
            Next
        End If
        
    End If
    
    QuickSort = arr
End Function












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
