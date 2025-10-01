Public Class InvtPartObj
    Dim oPrtName As String
    Dim oFilePath As String
    Dim nPrtName As String
    Dim nFilePath As String
    Dim oPrtDoc As Inventor.PartDocument
    Dim oFileName As String
    Dim nFileName As String
    Dim tNode As New TreeNode

    Property OriginalName As String
        Get
            Return oPrtName
        End Get
        Set(value As String)
            oPrtName = value
        End Set
    End Property

    Property OriginalFullFileName As String
        Get
            Return oFileName
        End Get
        Set(value As String)
            oFileName = value
        End Set
    End Property

    Property OriginalFilePath As String
        Get
            Return oFilePath
        End Get
        Set(value As String)
            oFilePath = value
        End Set
    End Property

    Property NewName As String
        Get
            Return nPrtName
        End Get
        Set(value As String)
            nPrtName = value
            NewFileName = nPrtName & ".ipt"
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

    Property OriginalPartDocument As Inventor.PartDocument
        Get
            Return oPrtDoc
        End Get
        Set(value As Inventor.PartDocument)
            oPrtDoc = value
        End Set
    End Property

    Property NewFileName As String
        Get
            Return nFileName
        End Get
        Set(value As String)
            nFileName = value
        End Set
    End Property

    Property NewTreeNode As TreeNode
        Get
            Return tNode
        End Get
        Set(value As TreeNode)
            tNode = value
        End Set
    End Property

    ReadOnly Property NewFullFileName As String
        Get
            Return nFilePath & nFileName
        End Get
    End Property


End Class
