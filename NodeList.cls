VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private root As Node
Private size As Long

Property Get Length() As Long
    Length = size
End Property

Sub Push(value As Variant)
    If root Is Nothing Then
        Set root = New Node
            root.data = value
    Else
    
        Dim n As New Node
            n.data = value
        Set n.next_ = root
        Set root = n
        
    End If
    
    size = size + 1
End Sub

Function Pop() As Variant

    If root Is Nothing Then
        Pop = Empty
    Else
        Pop = root.data
        
        If root.next_ Is Nothing Then
            Set root = Nothing
            size = 0
        Else
            Set root = root.next_
            size = size - 1
        End If
        
    End If
End Function

Sub Add(value As Variant)
    If root Is Nothing Then
        Set root = New Node
            root.data = value
    Else
    
        Dim current As Node
        Set current = root
    
        While Not (current.next_ Is Nothing)
            Set current = current.next_
        Wend
        
        Set current.next_ = New Node
            current.next_.data = value
    
    End If
    
    size = size + 1
    
End Sub

Sub List()
    If Not (root Is Nothing) Then
        
        Dim current As Node
        Set current = root
    
        While Not (current Is Nothing)
            JSON.Log current.data
            Set current = current.next_
        Wend
    End If
End Sub


