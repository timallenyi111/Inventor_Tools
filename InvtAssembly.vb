Imports System.Windows
Imports Inventor

Friend Class InvtAssembly
    Private prtList As List(Of InvtPart)
    Private subAsyList As List(Of InvtAssembly)
    Private subFrameList As List(Of InvtFrame)
    Private oAsyName As String
    Private _nAsyName As String
    Private oFullFileName As String
    Private _nFullFileName As String
    Private _nRootDirectory As String
    Private oAsmDoc As AssemblyDocument
    Private _treeNode As TreeNode
    Private oCompOcc As ComponentOccurrence
    Private _subType As String
    Private hltSet As HighlightSet
    Private ReadOnly duplicateOccurrenceList As List(Of ComponentOccurrence)
    Private ReadOnly _duplicateOccIndexList As List(Of Integer)
    Private _copyEnabled As Boolean = True
    Private _occurrenceIndex As Integer
    Private _containsFrame As Boolean = False

    ''' <summary>
    ''' automatically sets up the original file name, assembly name, and stores the assembly document. 
    ''' If an occurrence is passed it will also store that, but if not it assumes this is the root assembly and there won't be an occurrence associated with it.
    ''' </summary>
    ''' <param name="Asydoc"></param>
    ''' <param name="AsyOcc"></param>
    Public Sub New(ByRef Asydoc As Inventor.AssemblyDocument, ByRef occurrenceIndex As Integer, ByRef nRootDirectory As String,
                   Optional ByRef AsyOcc As ComponentOccurrence = Nothing)
        oFullFileName = Asydoc.FullFileName
        oAsmDoc = Asydoc
        oAsyName = GetAssemblyName(oFullFileName)
        _occurrenceIndex = occurrenceIndex
        _nRootDirectory = nRootDirectory

        If AsyOcc Is Nothing Then
            'This is a root assembly so it won't have an occurrence associated with it            
        Else

            oCompOcc = AsyOcc
        End If

        'initialize all of the applicable lists
        prtList = New List(Of InvtPart)
        subAsyList = New List(Of InvtAssembly)
        subFrameList = New List(Of InvtFrame)
        duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
        _duplicateOccIndexList = New List(Of Integer)
    End Sub


#Region "Private Functions"
    ''' <summary>
    ''' updates the tree node text with the new name
    ''' </summary>
    Private Sub NameChange()
        Try
            _nFullFileName = _nRootDirectory & _nAsyName & ".iam"

            ' rename the treeview node if the treeview has already been set up
            If _treeNode Is Nothing Then
                'the tree node hasn't been set up yet, so we don't need to worry about changing the name there yet since it will be set to the new name when it is created
            Else
                _treeNode.Text = _nAsyName
            End If

        Catch ex As Exception
            If _nRootDirectory Is Nothing Then
                'throw an exception
                Throw New Exception("Root directory is not set. Cannot change name until root directory is set.")

                'in the future we could automate this by looking for a root directory instead of throwing an error, but for now this will just force the user to set the root directory before changing any names which will ensure that the file paths are correct when they change the name.
            Else
                Throw New Exception("An error occurred while changing the name. " & ex.Message)
            End If
        End Try

    End Sub

    Private Function GetAssemblyName(ByRef fullFileName As String) As String
        Dim asyName As String = fullFileName.Substring(fullFileName.LastIndexOf("\") + 1)
        'now remove the .iam    
        asyName = asyName.Substring(0, asyName.Length - 4)
        Return asyName
    End Function

#End Region

#Region "Public Functions/Subs"
    ''' <summary>
    ''' Checks if the part is a duplicate occurrence in the assembly.
    ''' If it is then it adds the occurrence and occurrence index to the part object's duplicate occurrence and duplicate index list and returns false
    ''' </summary>
    ''' <param name="partOcc"></param>
    ''' <param name="occurrenceIndex"></param>
    ''' <returns></returns>
    Function CheckForDuplicatePart(ByRef partOcc As ComponentOccurrence, ByRef occurrenceIndex As Integer) As Boolean
        For Each prtObj As InvtPart In prtList
            If prtObj.OriginalFullFileName = partOcc.Definition.Document.FullFileName Then
                'this is a duplicate occurrence, so add it to the list of duplicates for this part object and return true
                prtObj.DuplicateOccurrences.Add(partOcc)
                prtObj.DuplicateOccurrenceIndexList.Add(occurrenceIndex)
                Return True
            End If
        Next
        Return False
    End Function

    Function CheckForDuplicateAssembly(ByRef asyOcc As ComponentOccurrence, ByRef occurrenceIndex As Integer) As Boolean
        For Each asyObj As InvtAssembly In subAsyList
            If asyObj.OriginalFullFileName = asyOcc.Definition.Document.FullFileName Then
                'this is a duplicate occurrence, so add it to the list of duplicates for this assembly object and return true                
                asyObj.DuplicateOccurrences.Add(asyOcc)
                asyObj.DuplicateOccurrenceIndexList.Add(occurrenceIndex)
                Return True
            End If
        Next
        Return False
    End Function

    Sub AddPartToList(ByRef part As InvtPart)
        prtList.Add(part)
    End Sub

    Sub AddSubAssemblyToList(ByRef asy As InvtAssembly)
        subAsyList.Add(asy)
    End Sub

    Sub AddSubFrameToList(ByRef frame As InvtFrame)
        subFrameList.Add(frame)
        _containsFrame = True
    End Sub

    ''' <summary>
    ''' For changing the root directory after initial setup
    ''' </summary>
    ''' <param name="newRootDirectory"></param>
    Sub ChangeRootDirectory(ByRef newRootDirectory As String)
        _nRootDirectory = newRootDirectory
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' This name is based on the original full file name of the assembly document
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property OriginalName As String
        Get
            Return oAsyName
        End Get
    End Property

    ReadOnly Property OriginalFullFileName As String
        Get
            Return oFullFileName
        End Get
    End Property

    ReadOnly Property OriginalAsmDocument As AssemblyDocument
        Get
            Return oAsmDoc
        End Get
    End Property

    ReadOnly Property OriginalComponentOccurrence As ComponentOccurrence
        Get
            Return oCompOcc
        End Get
    End Property

    Property NewName As String
        Get
            Return _nAsyName
        End Get
        Set(value As String)
            _nAsyName = value
            NameChange()
        End Set
    End Property
    Property TreeNode As TreeNode
        Get
            Return _treeNode
        End Get
        Set(value As TreeNode)
            _treeNode = value
        End Set
    End Property

    ''' <summary>
    ''' The new full file name is the root directory of this assembly and the new name + ".iam"
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property NewFullFileName As String
        Get
            Return _nRootDirectory & _nAsyName & ".iam"
        End Get
    End Property
    ReadOnly Property NewRootDirectory As String
        Get
            Return _nRootDirectory
        End Get
    End Property

    Property SubType As String
        Get
            Return _subType
        End Get
        Set(value As String)
            _subType = value
        End Set
    End Property

    ReadOnly Property PartList As List(Of InvtPart)
        Get
            Return prtList
        End Get
    End Property

    ReadOnly Property AssemblyList As List(Of InvtAssembly)
        Get
            Return subAsyList
        End Get
    End Property

    ReadOnly Property FrameList As List(Of InvtFrame)
        Get
            Return subFrameList
        End Get
    End Property

    ReadOnly Property DuplicateOccurrences As List(Of Inventor.ComponentOccurrence)
        Get
            Return duplicateOccurrenceList
        End Get
    End Property

    ReadOnly Property HighlightSet As HighlightSet
        Get
            Return hltSet
        End Get
    End Property

    Property CopyEnabled As Boolean
        Get
            Return _copyEnabled
        End Get
        Set(value As Boolean)
            _copyEnabled = value
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

    ReadOnly Property DuplicateOccurrenceIndexList As List(Of Integer)
        Get
            Return _duplicateOccIndexList
        End Get
    End Property

    ReadOnly Property ContainsFrame As Boolean
        Get
            Return _containsFrame
        End Get
    End Property

#End Region
End Class
