Public Class Component_Modification
    Private Sub Component_Modification_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TB_ModCompName.Text = AssemblyCopyToolForm.doubleClickNode.Text
    End Sub

    Private Sub Button_CompModAccept_Click(sender As Object, e As EventArgs) Handles Button_CompModAccept.Click
        AssemblyCopyToolForm.doubleClickNode.Text = TB_ModCompName.Text
        If CB_DoNotCopy.Checked Then
            AssemblyCopyToolForm.DontCopyNode()
        End If
        Me.Close()
    End Sub
End Class