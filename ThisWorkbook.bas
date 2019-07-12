VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ExportCode()
    
    Dim prj As VBProject
    Dim cmp As VBComponent
    Dim fso As New FileSystemObject
    Dim path As String
    Dim ext As String
    
    Set prj = ThisWorkbook.VBProject
    
    path = ThisWorkbook.path & "\VBA modules\"
    If Not fso.FolderExists(path) Then fso.CreateFolder path
    
    
    For Each cmp In prj.VBComponents
        ext = ".bas"
        
        Select Case cmp.Type
            Case vbext_ComponentType.vbext_ct_StdModule: ext = ".bas"
            Case vbext_ComponentType.vbext_ct_ClassModule: ext = ".cls"
        End Select
        
        cmp.Export path & cmp.name & ext
    Next
    
End Sub
