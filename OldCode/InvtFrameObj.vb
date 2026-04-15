Imports Inventor

Friend Class InvtFrameObj
    ' Initialized backing fields to safe defaults
    Private oFrmName As String = ""
    Private oFullFileName As String = ""
    Private oAsmDoc As PartDocument = Nothing
    Private oTreeNode As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode()
    Private oCompOcc As ComponentOccurrence = Nothing

    Private nFrmName As String = ""
    Private nTreeNode As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode()
    Private nFullFileName As String = ""
    Private nRootDirectory As String = ""



#Region "Properties"
    ReadOnly Property OriginalName As String
        Get
            Return oFrmName
        End Get
    End Property

    ReadOnly Property OriginalFullFileName As String
        Get
            Return oFullFileName
        End Get
    End Property

    ReadOnly Property OriginalPartDocument As PartDocument
        Get
            Return oAsmDoc
        End Get
    End Property

    ReadOnly Property OriginalTreeNode As TreeNode
        Get
            Return oTreeNode
        End Get
    End Property

    ReadOnly Property OriginalComponentOccurence As ComponentOccurrence
        Get
            Return oCompOcc
        End Get
    End Property

    ReadOnly Property NewName As String
        Get
            Return nFrmName
        End Get
    End Property
    ReadOnly Property NewTreeNode As TreeNode
        Get
            Return nTreeNode
        End Get
    End Property

    ReadOnly Property NewFullFileName As String
        Get
            Return nFullFileName
        End Get
    End Property
    ReadOnly Property NewRootDirectory As String
        Get
            Return nRootDirectory
        End Get
    End Property

#End Region
End Class
