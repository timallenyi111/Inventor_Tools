Friend Class InvtPartObj
    Private oPrtName As String
    Private oFilePath As String
    Private nPrtName As String
    Private nFilePath As String
    Private oPrtDoc As Inventor.PartDocument
    Private oFullFileName As String
    'Private nFileName As String
    Private tNode As New TreeNode
    Private oPartNumber As String
    Private nPartNumber As String
    Private oPartOcc As Inventor.ComponentOccurrence
    Private oTreeNode As TreeNode
    Private nTreeNode As TreeNode
    Private duplicateOccurrenceList As List(Of Inventor.ComponentOccurrence)
    Private _duplicateOccurrenceIndexList As List(Of Integer)
    Private _subType As String
    Private _enableCopy As Boolean = True
    Private _occurrenceIndex As Integer

#Region "Properties"
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

    ReadOnly Property OriginalComponentOccurrence As Inventor.ComponentOccurrence
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

    Property CopyEnabled As Boolean
        Get
            Return _enableCopy
        End Get
        Set(value As Boolean)
            _enableCopy = value
        End Set
    End Property

    Property OccurrenceIndex As Integer
        Get
            Return _occurrenceIndex
        End Get
        Set(value As Integer)
            _occurrenceIndex = value
        End Set
    End Property

    Property DuplicateOccurrenceIndexList As List(Of Integer)
        Get
            Return _duplicateOccurrenceIndexList
        End Get
        Set(value As List(Of Integer))
            _duplicateOccurrenceIndexList = value
        End Set
    End Property

#End Region

    Sub InitialSetup(ByRef PartOcc As Inventor.ComponentOccurrence, ByRef rootDirectory As String,
                     ByRef ParentAssemblyNewNode As TreeNode, ByRef contentCenterPath As String)

        Try
            oPrtDoc = PartOcc.Definition.Document
            oFullFileName = oPrtDoc.FullFileName
        Catch comEx As System.Runtime.InteropServices.COMException
            Debug.WriteLine("Could not get Definition.Document for occurrence: " & PartOcc.Name & " ⇒ " & comEx.Message)
            ' Try to get a referenced file path as a fallback
            Try
                Dim rfd = PartOcc.ReferencedFileDescriptor
                If rfd IsNot Nothing AndAlso rfd.ReferencedFile IsNot Nothing Then
                    oFullFileName = rfd.ReferencedFile.FullFileName
                Else
                    ' Last-resort: mark as content-center or missing; use occurrence name so logic continues
                    oFullFileName = PartOcc.Name
                    _subType = "Unknown/ContentCenterOrMissing"
                End If
            Catch ex As Exception
                Debug.WriteLine("Fallback failed: " & ex.Message)
                oFullFileName = PartOcc.Name
                _subType = "Unknown/ContentCenterOrMissing"
            End Try
        End Try
        oPrtName = GetPartName(oFullFileName)
        'check if this is a content center part
        If oFullFileName.ToLower.StartsWith(contentCenterPath.ToLower) Then
            _subType = "Content Center Part"
            nTreeNode = ParentAssemblyNewNode.Nodes.Add(oPrtName)
            Debug.WriteLine("Content Center Part Found: " & oPrtName)
            duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
            _duplicateOccurrenceIndexList = New List(Of Integer)
        Else
            nPrtName = oPrtName ' For now.... we aren't changing the new names by default
            nFilePath = rootDirectory
            nTreeNode = ParentAssemblyNewNode.Nodes.Add(nPrtName)
            duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
            _duplicateOccurrenceIndexList = New List(Of Integer)
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
        If NewTreeNode.ForeColor = System.Drawing.Color.Red Then
            CopyEnabled = False
        Else
            nPrtName = nTreeNode.Text
            'Content Center Parts retain their original file path
            If SubType IsNot "Content Center Part" Then
                nFilePath = rootDirectory
            End If
        End If

    End Sub

    Sub AddDuplicateOccurrence(ByRef dupOcc As Inventor.ComponentOccurrence, occIndex As Integer)
        duplicateOccurrenceList.Add(dupOcc)
        _duplicateOccurrenceIndexList.Add(occIndex)
    End Sub
End Class
