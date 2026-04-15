Imports System.Windows
Imports Inventor

Friend Class InvtAssembly
    Private _partList As List(Of InvtPart)
    Private _subAsyList As List(Of InvtAssembly)
    Private _subFrameList As List(Of InvtFrame)
    Private _oAsyName As String
    Private _nAsyName As String
    Private _oFullFileName As String
    Private _nFullFileName As String
    Private _nRootDirectory As String
    Private _oAsmDoc As AssemblyDocument
    Private _treeNode As TreeNode
    Private _oCompOcc As ComponentOccurrence
    Private _subType As String
    'Private _hltSet As HighlightSet
    Private ReadOnly duplicateOccurrenceList As List(Of ComponentOccurrence)
    Private ReadOnly _duplicateOccIndexList As List(Of Integer)
    Private _copyEnabled As Boolean = True
    Private _occurrenceIndex As Integer
    Private _containsFrame As Boolean = False
    Private _isBoltedConnection As Boolean = False
    Private _parentName As String

    ''' <summary>
    ''' automatically sets up the original file name, assembly name, stores the assembly document. 
    ''' If an occurrence is passed it will also store that, but if not it assumes this is the root assembly and there won't be an occurrence associated with it | 
    ''' Checks if it is a bolted connection and if so performs the propery root directory operations    
    ''' </summary>
    ''' <param name="Asydoc"></param>
    ''' <param name="AsyOcc"></param>
    Public Sub New(ByRef Asydoc As Inventor.AssemblyDocument, ByRef occurrenceIndex As Integer, ByRef nRootDirectory As String,
                   Optional ByRef AsyOcc As ComponentOccurrence = Nothing)
        _oFullFileName = Asydoc.FullFileName
        _oAsmDoc = Asydoc
        _oAsyName = GetAssemblyName(_oFullFileName)
        _occurrenceIndex = occurrenceIndex
        _nRootDirectory = nRootDirectory

        If AsyOcc Is Nothing Then
            'This is a root assembly so it won't have an occurrence associated with it            
        Else
            _oCompOcc = AsyOcc
            CheckIfBoltedConnection()
        End If

        'initialize all of the applicable lists
        _partList = New List(Of InvtPart)
        _subAsyList = New List(Of InvtAssembly)
        _subFrameList = New List(Of InvtFrame)
        duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
        _duplicateOccIndexList = New List(Of Integer)
    End Sub


#Region "Private Subs/Functions"
    ''' <summary>
    ''' updates the tree node text with the new name
    ''' </summary>
    Private Sub NameChange()
        Try

            ' rename the treeview node if the treeview has already been set up
            If _treeNode Is Nothing Then
                'the tree node hasn't been set up yet, so we don't need to worry about changing the name there yet since it will be set to the new name when it is created
            Else
                _treeNode.Text = _nAsyName
                If _treeNode.Parent Is Nothing Then
                    'this is the root so when the name changes, so does the root directory
                    Dim lastIndex As Integer = _nRootDirectory.LastIndexOf("\")
                    _nRootDirectory = _nRootDirectory.Substring(0, lastIndex)
                    lastIndex = _nRootDirectory.LastIndexOf("\")
                    _nRootDirectory = _nRootDirectory.Substring(0, lastIndex + 1) & _nAsyName & "\"
                    UpdateSubCompRootDir()
                End If
            End If

            _nFullFileName = _nRootDirectory & _nAsyName & ".iam"

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

    ''' <summary>
    ''' Checks if the occurrence is a bolted connection and sets the _isBoltedConnection flag | 
    ''' Also performs the necessary root directory 
    ''' </summary>
    ''' <returns></returns>
    Private Sub CheckIfBoltedConnection()
        If _oCompOcc.AttributeSets.Count > 0 Then
            For Each attSet As AttributeSet In _oCompOcc.AttributeSets
                If attSet.Name = "FDesign" Then
                    For Each atri As Inventor.Attribute In attSet
                        If VarType(atri.Value) = vbString Then
                            Dim atriValue As String = atri.Value
                            If atriValue.IndexOf("CABoltCon") >= 0 Then
                                _isBoltedConnection = True
                                Exit Sub
                            End If
                        End If
                    Next
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' Used for passing the root directory down to sub components
    ''' </summary>
    ''' <param name="newRootDirectory"></param>
    Private Sub UpdateSubCompRootDir()
        For Each part As InvtPart In _partList
            part.NewRootDirectory = _nRootDirectory
        Next
        For Each subAsm As InvtAssembly In _subAsyList
            subAsm.NewRootDirectory = _nRootDirectory
        Next
        For Each subFrm As InvtFrame In _subFrameList
            subFrm.NewRootDirectory = _nRootDirectory
        Next
    End Sub

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
        For Each prtObj As InvtPart In _partList
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
        For Each asyObj As InvtAssembly In _subAsyList
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
        _partList.Add(part)
    End Sub

    Sub AddSubAssemblyToList(ByRef asy As InvtAssembly)
        _subAsyList.Add(asy)
    End Sub

    Sub AddSubFrameToList(ByRef frame As InvtFrame)
        _subFrameList.Add(frame)
        _containsFrame = True
    End Sub



#End Region

#Region "Properties"

    ''' <summary>
    ''' This name is based on the original full file name of the assembly document
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property OriginalName As String
        Get
            Return _oAsyName
        End Get
    End Property

    ReadOnly Property OriginalFullFileName As String
        Get
            Return _oFullFileName
        End Get
    End Property

    ReadOnly Property OriginalAsmDocument As AssemblyDocument
        Get
            Return _oAsmDoc
        End Get
    End Property

    ReadOnly Property OriginalComponentOccurrence As ComponentOccurrence
        Get
            Return _oCompOcc
        End Get
    End Property

    Property NewName As String
        Get
            If _treeNode Is Nothing Then
                Return _nAsyName
            Else
                Return _treeNode.Text
            End If
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
            If _isBoltedConnection Then
                _treeNode.ForeColor = System.Drawing.Color.Gray
            End If
        End Set
    End Property

    ''' <summary>
    ''' The new full file name is the root directory of this assembly and the new name + ".iam"
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property NewFullFileName As String
        Get
            If _isBoltedConnection Then
                Return _nRootDirectory & _treeNode.Parent.Text & "\Design Accelerator\" & _nAsyName & ".iam"
            Else
                Return _nRootDirectory & _nAsyName & ".iam"
            End If
        End Get
    End Property
    Property NewRootDirectory As String
        Get
            Return _nRootDirectory
        End Get
        Set(value As String)
            _nRootDirectory = value
            UpdateSubCompRootDir()
        End Set
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
            Return _partList
        End Get
    End Property

    ReadOnly Property AssemblyList As List(Of InvtAssembly)
        Get
            Return _subAsyList
        End Get
    End Property

    ReadOnly Property FrameList As List(Of InvtFrame)
        Get
            Return _subFrameList
        End Get
    End Property

    ReadOnly Property DuplicateOccurrences As List(Of Inventor.ComponentOccurrence)
        Get
            Return duplicateOccurrenceList
        End Get
    End Property

    'ReadOnly Property HighlightSet As HighlightSet
    '    Get
    '        Return _hltSet
    '    End Get
    'End Property

    Property CopyEnabled As Boolean
        Get
            If _treeNode.ForeColor = System.Drawing.Color.Red Then
                _copyEnabled = False
            End If
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
