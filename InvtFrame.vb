Imports System.Windows
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Inventor

''' <summary>
''' This is a class that represents a frame generator assembly in inventor.
''' It uses an InvtAssembly Object to store the assembly-level information but then adds additional functions and properties
''' to handle frame specific tasks
''' </summary>
Friend Class InvtFrame

    'the core of this class uses an InvtAssembly
    Private ReadOnly _frameAssemblyObject As InvtAssembly

    Private _oSkeletonID As String
    Private _nSkeletonID As String
    Private _parentAssemblyName As String 'needed for making root directory

    ''' <summary>
    ''' By sending an InvtAssemblyObject when we create the frame object, we can use the setup that is already done in the InvtAssembly class to store all of the assembly level information and then just add on the frame specific information and functions in this class.
    ''' then the frame specefic setup will be done automatically.
    ''' </summary>
    ''' <param name="InvtAssemblyObject"></param>
    Public Sub New(ByRef InvtAssemblyObject As InvtAssembly, ByRef parentAssemblyName As String)
        'setup a sub assembly object to represent the frame assembly
        _frameAssemblyObject = InvtAssemblyObject
        _parentAssemblyName = parentAssemblyName
        'now we do the frame specific setup

        SetSkeletonIDs(_frameAssemblyObject.OriginalComponentOccurrence)

        GenerateNewFrameName()

        'change the file paths for the frame along with its sub components to be stored in the "Frame" directory
        ChangeRootDirectory(_frameAssemblyObject.NewRootDirectory)

    End Sub

    '******properties from InvtAssembly Object
#Region "Properties from InvtAssembly"

    ''' <summary>
    ''' This name is based on the original full file name of the assembly document
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property OriginalName As String
        Get
            Return _frameAssemblyObject.OriginalName
        End Get
    End Property

    ReadOnly Property OriginalFullFileName As String
        Get
            Return _frameAssemblyObject.OriginalFullFileName
        End Get
    End Property

    ReadOnly Property OriginalAsmDocument As AssemblyDocument
        Get
            Return _frameAssemblyObject.OriginalAsmDocument
        End Get
    End Property

    ReadOnly Property OriginalComponentOccurrence As ComponentOccurrence
        Get
            Return _frameAssemblyObject.OriginalComponentOccurrence
        End Get
    End Property

    Property NewName As String
        Get
            Return _frameAssemblyObject.NewName
        End Get
        Set(value As String)
            _frameAssemblyObject.NewName = value
        End Set
    End Property
    Property TreeNode As TreeNode
        Get
            Return _frameAssemblyObject.TreeNode
        End Get
        Set(value As TreeNode)
            _frameAssemblyObject.TreeNode = value
        End Set
    End Property

    ReadOnly Property NewFullFileName As String
        Get
            Return _frameAssemblyObject.NewFullFileName
        End Get
    End Property
    ReadOnly Property NewRootDirectory As String
        Get
            Return _frameAssemblyObject.NewRootDirectory
        End Get
    End Property

    Property SubType As String
        Get
            Return _frameAssemblyObject.SubType
        End Get
        Set(value As String)
            _frameAssemblyObject.SubType = value
        End Set
    End Property

    ReadOnly Property PartList As List(Of InvtPart)
        Get
            Return _frameAssemblyObject.PartList
        End Get
    End Property

    ReadOnly Property AssemblyList As List(Of InvtAssembly)
        Get
            Return _frameAssemblyObject.AssemblyList
        End Get
    End Property

    ReadOnly Property FrameList As List(Of InvtFrame)
        Get
            Return _frameAssemblyObject.FrameList
        End Get
    End Property

    ReadOnly Property DuplicateOccurrences As List(Of Inventor.ComponentOccurrence)
        Get
            Return _frameAssemblyObject.DuplicateOccurrences
        End Get
    End Property

    ReadOnly Property HighlightSet As HighlightSet
        Get
            Return _frameAssemblyObject.HighlightSet
        End Get
    End Property

    Property CopyEnabled As Boolean
        Get
            Return _frameAssemblyObject.CopyEnabled
        End Get
        Set(value As Boolean)
            _frameAssemblyObject.CopyEnabled = value
        End Set
    End Property

    Property OccurrenceIndex As Integer
        Get
            Return _frameAssemblyObject.OccurrenceIndex
        End Get
        Set(value As Integer)
            _frameAssemblyObject.OccurrenceIndex = value
        End Set
    End Property

    ReadOnly Property DuplicateOccurrenceIndexList As List(Of Integer)
        Get
            Return _frameAssemblyObject.DuplicateOccurrenceIndexList
        End Get
    End Property

    ReadOnly Property ContainsFrame As Boolean
        Get
            Return _frameAssemblyObject.ContainsFrame
        End Get
    End Property

#End Region


    '******Frame Specific Properties
#Region "Frame Specific Properties"

    ReadOnly Property OriginalSkeletonID As String
        Get
            Return _oSkeletonID
        End Get
    End Property

    ReadOnly Property NewSkeletonID As String
        Get
            Return _nSkeletonID
        End Get
    End Property

#End Region


    '***** Public Functions
#Region "Public Functions"
    Sub AddPartToList(ByRef part As InvtPart)
        'first we have to set the part root directory to the frame
        part.ChangeRootDirectory(_frameAssemblyObject.NewRootDirectory)
        _frameAssemblyObject.AddPartToList(part)
    End Sub

    Sub AddSubAssemblyToList(ByRef asy As InvtAssembly)
        asy.ChangeRootDirectory(_frameAssemblyObject.NewRootDirectory)
        _frameAssemblyObject.AddSubAssemblyToList(asy)
    End Sub

    Sub AddSubFrameToList(ByRef frame As InvtFrame)
        frame.ChangeRootDirectory(_frameAssemblyObject.NewRootDirectory)
        _frameAssemblyObject.AddSubFrameToList(frame)
    End Sub

    ''' <summary>
    ''' For changing the root directory of a frame to be stored in the "Frame" subdirectory after initial setup. This function also changes the root directory for all of the parts in the frame to be stored in the same "Frame" subdirectory. We need to do this because the frame assembly and all of its components need to be stored in the same location in order for the assembly copying and updating functions to work correctly.
    ''' </summary>
    ''' <param name="inputRootDirectory"></param>
    Sub ChangeRootDirectory(ByRef inputRootDirectory As String)
        Dim newRootDirectory As String = inputRootDirectory & _parentAssemblyName & "\Frame\"
        _frameAssemblyObject.ChangeRootDirectory(newRootDirectory)
        For Each part As InvtPart In PartList
            part.ChangeRootDirectory(newRootDirectory)
        Next
        For Each asy As InvtAssembly In AssemblyList
            asy.ChangeRootDirectory(newRootDirectory)
        Next
        For Each frame As InvtFrame In FrameList
            frame.ChangeRootDirectory(newRootDirectory)
        Next

    End Sub

#End Region

    '******Private Functions
#Region "Frame Specific Functions"

    ''' <summary>
    ''' Sets the original and new skeleton IDs for the frame assembly. The skeleton IDs are stored in the frame assembly occurrence attributes under "Frame.Skeletons". We need to get the original skeleton ID from there and then generate a new skeleton ID to use when we update the frame assembly occurrence attributes during the copy process.
    ''' </summary>
    ''' <param name="frmOcc"></param>
    Private Sub SetSkeletonIDs(ByRef frmOcc As Inventor.ComponentOccurrence)
        'search through the frame occurrence definition attributes for the skeleton id and store it as the original skeleton id,
        'then generate a new skeleton id and store it
        For Each attSet As Inventor.AttributeSet In frmOcc.Definition.AttributeSets
            For Each atri As Inventor.Attribute In attSet
                If atri.Name = "Frame.Skeletons" Then
                    Dim atriVal As String = atri.Value
                    'the integer value for the start of the skeleton id in the atribute value string
                    Dim skelIDStartInt As Integer = GetSkelIdStartInt(atriVal)
                    'the integer value for the end of the skeleton id in the atribute value string
                    Dim skelIDEndInt As Integer = GetSkelIdEndInt(atriVal, skelIDStartInt)
                    _oSkeletonID = atriVal.Substring(skelIDStartInt, skelIDEndInt - skelIDStartInt)
                    _nSkeletonID = GenerateNewID(_oSkeletonID)
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Get the integer for the start of the skeleton id in the frame assembly attribute value
    ''' </summary>
    ''' <param name="atri"></param>
    ''' <returns></returns>
    Private Function GetSkelIdStartInt(ByVal atri As String) As Integer
        Dim skelIDStart = InStr(atri, "SkeletonID")
        Dim skelId As String = atri.Substring(skelIDStart)
        skelIDStart = skelIDStart + InStr(skelId, """")
        Return skelIDStart
    End Function

    Private Function GetPathIdStartInt(ByVal atri As String) As Integer
        Dim pathIDStart = InStr(atri, "PathID")
        Dim pathId As String = atri.Substring(pathIDStart)
        pathIDStart = pathIDStart + InStr(pathId, """")
        Return pathIDStart
    End Function

    ''' <summary>
    ''' Gets the integer for the end of the skeleton id in the frame assembly attribute value
    ''' </summary>
    ''' <param name="atri"></param>
    ''' <param name="skelIdStart"></param>
    ''' <returns></returns>
    Private Function GetSkelIdEndInt(ByVal atri As String, ByVal skelIdStart As Integer) As Integer
        Dim skelId As String = atri.Substring(skelIdStart)
        Dim skelIdEnd As Integer = skelIdStart + InStr(skelId, """") - 1
        Return skelIdEnd
    End Function

    Private Function GetPathIdEndInt(ByVal atri As String, ByVal pathIdStart As Integer) As Integer
        Dim pathId As String = atri.Substring(pathIdStart)
        Dim pathIdEnd As Integer = pathIdStart + InStr(pathId, """") - 1
        Return pathIdEnd
    End Function

    ''' <summary>
    ''' Replaces everything after the final "-" in the original skeleton id with random integers
    ''' </summary>
    ''' <param name="oSkelId"></param>
    ''' <returns></returns>
    Private Function GenerateNewID(ByVal oSkelId As String) As String
        Dim newSkelIdEnd As String = oSkelId.Substring(oSkelId.LastIndexOf("-") + 1)
        Debug.WriteLine("SkeletonID End: " & newSkelIdEnd)
        Dim i As Integer = 0
        Dim rnd As New Random
        While i < newSkelIdEnd.Length
            Dim newInt As Integer = rnd.Next(0, 9)
            Dim newChar As String = newInt.ToString
            newSkelIdEnd = newSkelIdEnd.Substring(0, i) & newChar & newSkelIdEnd.Substring(i + 1)
            i += 1
        End While
        Dim newSkelId As String = oSkelId.Substring(0, oSkelId.LastIndexOf("-") + 1) & newSkelIdEnd
        Return newSkelId
    End Function

    ''' <summary>
    ''' Sets the new frame name automatically with random integers.
    ''' This is necessary because the frame generator assemblies need to have unique names. ''' </summary>
    ''' <param name="frameName"></param>
    Private Sub GenerateNewFrameName()
        Dim nFrameName As String = "Frame_"
        Dim rnd As New Random
        Dim count = 0
        While count < 13
            nFrameName = nFrameName + rnd.Next(0, 9).ToString
            count += 1
        End While
        _frameAssemblyObject.NewName = nFrameName
    End Sub

    Private Sub SetFrameRootDirectory()
        'the root directory that was set for the core assembly object during setup
        Dim assemblyRootDirectory = _frameAssemblyObject.NewRootDirectory
        'frames are stored in a subdirectory based on their parent assembly name
        _frameAssemblyObject.ChangeRootDirectory(assemblyRootDirectory & _parentAssemblyName & "\Frame\")
    End Sub

#End Region
    '******************************

End Class
