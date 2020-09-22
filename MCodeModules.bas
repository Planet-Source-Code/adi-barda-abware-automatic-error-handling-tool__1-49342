Attribute VB_Name = "MCodeModules"
Option Explicit

Private m_objTrv As abTreeView
Private m_iKey As Long

Public Sub SetTreeview(ByRef objTrv As abTreeView)
    Set m_objTrv = objTrv
End Sub


Public Sub ClearTreeView()

    With m_objTrv.SourceTreeView
        .Nodes.Clear
    End With

End Sub

Public Sub AddCodeModule(ByVal sParent As String, _
                         ByVal sKey As String, _
                         ByVal sText As String, _
                         Optional ByVal sTag As String = "", _
                         Optional ByVal sImage As String = "ROOT")

    Dim nodx As MSComctlLib.Node
    
    With m_objTrv.SourceTreeView
        Set nodx = .Nodes.Add(sParent, tvwChild, sKey, sText)
        nodx.Tag = sTag
        nodx.Image = sImage
        nodx.Checked = True
    End With
    
End Sub

Public Sub SetModulesStruct()

    Dim nodx As MSComctlLib.Node
    
    With m_objTrv.SourceTreeView
    
        .Indentation = 1
        .Checkboxes = True
        
        Set nodx = .Nodes.Add(, , "Project", "Project")
        nodx.Tag = ""
        nodx.Image = "ROOT"
        nodx.Expanded = True
        
        AddCodeModule "Project", "Forms", "Forms", , "FOLDER"
        AddCodeModule "Project", "Modules", "Modules", , "FOLDER"
        AddCodeModule "Project", "Classes", "Classes", , "FOLDER"
        AddCodeModule "Project", "User controls", "User controls", , "FOLDER"
        
        
    End With
    
End Sub

Public Function FileInList(ByVal sFileName As String) As Boolean

    'check if the specified file is allready in the list
    

    Dim i           As Long

    FileInList = False

    With m_objTrv.SourceTreeView
        For i = 1 To .Nodes.Count
            FileInList = (sFileName = .Nodes.Item(i).Tag)
            If FileInList Then Exit For
        Next i
    End With
    
Exit_Proc:
    Exit Function


End Function

Public Function GetNextKey() As String

    m_iKey = m_iKey + 1
    GetNextKey = "K" & m_iKey
    
End Function
