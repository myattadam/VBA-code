Attribute VB_Name = "JSON"
Option Explicit
Option Base 1

Sub Save(filename As String, Entity As Variant)
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fs.OpenTextFile(ActiveWorkbook.path & "\" & filename, ForWriting, True)
    ts.Write toString(Entity)
    ts.Close
End Sub

Sub Log(Entity As Variant)
    Debug.Print toString(Entity)
End Sub

Function toString(Entity As Variant) As String
    
    Dim index As Long
    Dim s As String
    
    If IsArray(Entity) Then
        s = s & "["

        For index = LBound(Entity) To UBound(Entity)
            s = s & toString(Entity(index))
            If index < UBound(Entity) Then s = s & ","
        Next

        s = s & "]"
    
    Else
    
        Select Case TypeName(Entity)
            Case "Empty"
                s = s & "null"
            
            Case "Integer", "Long", "Single", "Double"
                s = s & Entity
                
            Case "Boolean"
                s = s & """" & Entity & """"
                
            Case "String"
                s = s & """" & Entity & """"
                
            Case "Dictionary"
                s = s & "{"
                
                Dim keylist As Variant
                keylist = Entity.Keys
        
                For index = LBound(keylist) To UBound(keylist)
                    s = s & """" & keylist(index) & """:"
                    s = s & toString(Entity(keylist(index)))
                    If index < UBound(keylist) Then s = s & ","
                Next
                
                s = s & "}"
                
            
            Case "Collection"
                s = s & "["
                
                For index = 1 To Entity.count
                    s = s & toString(Entity(index))
                    If index < Entity.count Then s = s & ","
                Next
                
                s = s & "]"
            
            Case Else
                s = s & """" & TypeName(Entity) & """"
            
        End Select
    End If
    
    toString = s
    
End Function
