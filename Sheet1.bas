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
