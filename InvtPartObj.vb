Public Class InvtPartObj
    Dim prtName As String

    Property Name As String
        Get
            Return prtName
        End Get
        Set(value As String)
            prtName = value
        End Set
    End Property

End Class
