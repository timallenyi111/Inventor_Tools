Public Class InvtPart
    Private oPrtName As String
    Private _nPartName As String
    Private _newFullFileName As String
    Private _nRootDirectory As String
    Private oPrtDoc As Inventor.PartDocument
    Private oFullFileName As String
    Private oPartNumber As String
    Private nPartNumber As String
    Private oPartOcc As Inventor.ComponentOccurrence
    Private _treeNode As TreeNode
    Private duplicateOccurrenceList As List(Of Inventor.ComponentOccurrence)
    Private _duplicateOccurrenceIndexList As List(Of Integer)
    Private _subType As String
    Private _enableCopy As Boolean = True
    Private _occurrenceIndex As Integer
    Private _isContentCenter As Boolean = False

    Public Sub New(ByRef prtDoc As Inventor.PartDocument, ByRef partOcc As Inventor.ComponentOccurrence,
                   ByRef occurrenceIndex As Integer, ByRef nRootdirectory As String, Optional ByRef isContentCenter As Boolean = False)

        oFullFileName = prtDoc.FullFileName
        _occurrenceIndex = occurrenceIndex
        _nRootDirectory = nRootdirectory
        If isContentCenter Then
            _isContentCenter = True
        Else
            oPrtDoc = prtDoc
            oPartOcc = partOcc
        End If

        oPrtName = GetPartName(oFullFileName)

        'initialize the list to hold duplicate occurrences
        duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
        _duplicateOccurrenceIndexList = New List(Of Integer)

    End Sub

#Region "Functions"
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

    Private Sub NameChange()
        _newFullFileName = _nRootDirectory & _nPartName & ".ipt"
        If _treeNode IsNot Nothing Then
            _treeNode.Text = _nPartName
        End If
    End Sub

    Public Sub ChangeRootDirectory(nRootDirectory)
        _nRootDirectory = nRootDirectory
    End Sub

#End Region

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
            Return _nPartName
        End Get
        Set(value As String)
            _nPartName = value
            NameChange()
        End Set
    End Property

    Property TreeNode As TreeNode
        Set(value As TreeNode)
            _treeNode = value
        End Set
        Get
            Return _treeNode
        End Get
    End Property

    ''' <summary>
    ''' NewFullFileName is "{RootDirectory}/{NewName}.ipt"
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property NewFullFileName As String
        Get
            Return _nRootDirectory & _nPartName & ".ipt"
        End Get
    End Property

    ReadOnly Property NewRootDirectory As String
        Get
            Return _nRootDirectory
        End Get
    End Property

    ReadOnly Property IsContentCenter As Boolean
        Get
            Return _isContentCenter
        End Get
    End Property

    ''' <summary>
    ''' list of all component occurrences that are duplicates
    ''' </summary>
    ''' <returns>List(Of Inventor.ComponentOccurrence)</returns>
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

End Class
