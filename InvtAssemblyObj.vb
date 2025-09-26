Public Class InvtAssemblyObj
    Dim compDefs As List(Of Inventor.ComponentDefinition)
    Dim asmDef As Inventor.AssemblyComponentDefinition
    Dim asmObjList As List(Of InvtAssemblyObj)
    Dim compNames As List(Of String)
    Dim asmNames As List(Of String)
    Dim prtNames As List(Of String)
    Dim assmName As String
    Dim assmFileName As String
    Dim assmFilePath As String

    Public Sub New()
        compDefs = New List(Of Inventor.ComponentDefinition)()
        asmObjList = New List(Of InvtAssemblyObj)()
        compNames = New List(Of String)()
        asmNames = New List(Of String)()
        prtNames = New List(Of String)()
    End Sub

    Property Name As String
        Get
            Return assmName
        End Get
        Set(value As String)
            assmName = value
        End Set
    End Property

    Property FileName As String
        Get
            Return assmFileName
        End Get
        Set(value As String)
            assmFileName = value
        End Set
    End Property

    Property ComponentDefinitions As List(Of Inventor.ComponentDefinition)
        Get
            Return compDefs
        End Get
        Set(value As List(Of Inventor.ComponentDefinition))
            compDefs = value
        End Set
    End Property

    Property FilePath As String
        Get
            Return assmFilePath
        End Get
        Set(value As String)
            assmFilePath = value
        End Set
    End Property

    Property ComponentNames As List(Of String)
        Get
            Return compNames
        End Get
        Set(value As List(Of String))
            compNames = value
        End Set
    End Property

    Property SubAssemblyObjectList As List(Of InvtAssemblyObj)
        Get
            Return asmObjList
        End Get
        Set(value As List(Of InvtAssemblyObj))
            asmObjList = value
        End Set
    End Property

    Property AssemblyNames As List(Of String)
        Get
            Return asmNames
        End Get
        Set(value As List(Of String))
            asmNames = value
        End Set
    End Property

    Property PartNames As List(Of String)
        Get
            Return prtNames
        End Get
        Set(value As List(Of String))
            prtNames = value
        End Set
    End Property

    Property AssemblyComponentDefinition As Inventor.AssemblyComponentDefinition
        Get
            Return asmDef
        End Get
        Set(value As Inventor.AssemblyComponentDefinition)
            asmDef = value
        End Set
    End Property

End Class
