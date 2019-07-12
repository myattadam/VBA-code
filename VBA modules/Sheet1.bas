VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub test()
    Dim records As New Dictionary
    
    Set records("A") = New Dictionary
        records("A") = "FOO"
        records("B") = "BAR"
    Set records("C") = New Dictionary
        records("C")("Item 1") = "FIZZ"
        records("C")("Item 2") = "BUZZ"
    
    JSON.Save "test.json", records
    
End Sub




Sub test2()
    Dim a As New ArrayList
    
    
    Dim v(100) As Variant
    Dim total As Double
    
    For i = 0 To 100
        v(i) = Rnd() * 1000
    Next
    
    a.Add v
    
    JSON.Log a.List
    
    a.Clear
    
    a.Add v, True
    
    JSON.Log a.List
    
    total = a.Calculate(sum)
    a.Calculate Divide, total
    a.Calculate Multiply, 100
    a.ApplyFormat "0.0"
    JSON.Log a.List

End Sub
