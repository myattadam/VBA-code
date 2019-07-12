Attribute VB_Name = "Math"
Option Explicit
Option Base 1

Public Const pX# = 28.3464555555556
Public Const TAU# = 6.28318530717959


' Proper Mod function
Function nMod(a, n)
    nMod = a - (n * Int(a / n))
End Function


' Returns an array length 2, containing the slope (1) and Y-intercept (2)
Function getTrend(ByVal pair As Variant) As Variant
    Const X As Long = 1
    Const Y As Long = 2

    Dim i As Long
    Dim lower As Long
    Dim upper As Long
    Dim size As Long
    Dim mean(1 To 2) As Double
    
    lower = LBound(pair)
    upper = UBound(pair)
    size = upper - lower + 1
    
    For i = lower To upper
        mean(X) = mean(X) + pair(i)(X)
        mean(Y) = mean(Y) + pair(i)(Y)
    Next
    
    mean(X) = mean(X) / size
    mean(Y) = mean(Y) / size

    Dim temp As Variant
    ReDim temp(LBound(pair) To UBound(pair))
    
    For i = lower To upper
        temp(i) = Array((pair(i)(X) - mean(X)) * (pair(i)(Y) - mean(Y)), _
                        (pair(i)(X) - mean(X)) ^ 2)
    Next
    
    Dim sum(1 To 2) As Variant
    
    For i = lower To upper
        sum(1) = sum(1) + temp(i)(1)
        sum(2) = sum(2) + temp(i)(2)
    Next
    
    sum(1) = sum(1) / size
    sum(2) = sum(2) / size
    
    Dim result(1 To 2) As Double
        result(1) = sum(1) / sum(2) ' Slope
        result(2) = mean(Y) - (mean(X) * result(1)) ' Y-intercept

    getTrend = result
End Function

