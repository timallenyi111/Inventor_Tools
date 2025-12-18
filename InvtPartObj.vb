Friend Class InvtPartObj
    Dim oPrtName As String
    Dim oFilePath As String
    Dim nPrtName As String
    Dim nFilePath As String
    Dim oPrtDoc As Inventor.PartDocument
    Dim oFullFileName As String
    'Dim nFileName As String
    Dim tNode As New TreeNode
    Dim oPartNumber As String
    Dim nPartNumber As String
    Dim oPartOcc As Inventor.ComponentOccurrence
    Dim oTreeNode As TreeNode
    Dim nTreeNode As TreeNode
    Dim duplicateOccurrenceList As List(Of Inventor.ComponentOccurrence)
    Dim _subType As String

    ReadOnly Property OriginalName As String
        Get
            Return oPrtName
        End Get
    End Property

    ReadOnly Property OriginalFullFileName As String
        Get
            Return oFullFileName
        End Get
    End Property

    ''' <summary>
    ''' returns the file path to the original part disk location including the final "\"
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property OriginalFilePath As String
        Get
            Return oFullFileName.Substring(0, oFullFileName.LastIndexOf("\") + 1)
        End Get
    End Property

    ReadOnly Property OriginalComponentOccurence As Inventor.ComponentOccurrence
        Get
            Return oPartOcc
        End Get
    End Property

    Property OriginalPartDocument As Inventor.PartDocument
        Get
            Return oPrtDoc
        End Get
        Set(value As Inventor.PartDocument)
            oPrtDoc = value
        End Set
    End Property

    Property NewName As String
        Get
            Return nPrtName
        End Get
        Set(value As String)
            nPrtName = value
        End Set
    End Property

    Property NewFilePath As String
        Get
            Return nFilePath
        End Get
        Set(value As String)
            nFilePath = value
        End Set
    End Property


    ReadOnly Property NewTreeNode As TreeNode
        Get
            Return nTreeNode
        End Get
    End Property

    ReadOnly Property NewFullFileName As String
        Get
            Return nFilePath & nPrtName & ".ipt"
        End Get
    End Property

    ReadOnly Property DuplicateOccurrences As List(Of Inventor.ComponentOccurrence)
        Get
            Return duplicateOccurrenceList
        End Get
    End Property

    ReadOnly Property SubType As String
        Get
            Return _subType
        End Get
    End Property

    Sub InitialSetup(ByRef PartOcc As Inventor.ComponentOccurrence, ByRef rootDirectory As String,
                     ByRef ParentAssemblyNewNode As TreeNode, ByRef contentCenterPath As String)

        oPartOcc = PartOcc
        oPrtDoc = PartOcc.Definition.Document
        oFullFileName = oPrtDoc.FullFileName
        oPrtName = GetPartName(oFullFileName)
        'check if this is a content center part
        If oFullFileName.ToLower.StartsWith(contentCenterPath.ToLower) Then
            _subType = "Content Center Part"
            nTreeNode = ParentAssemblyNewNode.Nodes.Add(oPrtName)
            Debug.WriteLine("Content Center Part Found: " & oPrtName)
            duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
        Else
            nPrtName = oPrtName ' For now.... we aren't changing the new names by default
            nFilePath = rootDirectory
            nTreeNode = ParentAssemblyNewNode.Nodes.Add(nPrtName)
            duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
        End If


    End Sub

    ''' <summary>
    ''' Returns a part name that is based on the original file name without the .ipt
    ''' </summary>
    ''' <param name="FullFileName"></param>
    ''' <returns></returns>
    Private Function GetPartName(ByRef FullFileName As String) As String
        Dim partName As String = FullFileName.Substring(FullFileName.LastIndexOf("\") + 1)
        'now remove the .ipt
        partName = partName.Substring(0, partName.Length - 4)
        Return partName
    End Function

    Sub UpdateNewProperties(ByVal rootDirectory As String)
        nPrtName = nTreeNode.Text
        'Content Center Parts retain their original file path
        If SubType IsNot "Content Center Part" Then
            nFilePath = rootDirectory
        End If
    End Sub

    Sub AddDuplicateOccurrence(ByRef dupOcc As Inventor.ComponentOccurrence)
        duplicateOccurrenceList.Add(dupOcc)
    End Sub
End Class
