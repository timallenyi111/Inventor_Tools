Public Class Component_Modification
    Private Sub Component_Modification_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim clickedNode As TreeNode = AssemblyCopyToolForm.doubleClickNode
        TB_ModCompName.Text = clickedNode.Text
        If clickedNode.ForeColor = System.Drawing.Color.Red Then
            CB_DoNotCopy.Checked = True
        Else
            CB_DoNotCopy.Checked = False
        End If
    End Sub

    Private Sub Button_CompModAccept_Click(sender As Object, e As EventArgs) Handles Button_CompModAccept.Click
        Dim clickedNode As TreeNode = AssemblyCopyToolForm.doubleClickNode
        clickedNode.Text = TB_ModCompName.Text
        If CB_DoNotCopy.Checked Then
            AssemblyCopyToolForm.DontCopyNode()
        End If

        If clickedNode.Parent Is Nothing Then
            'adjust the root directory of the root assembly if that is the node that has been changed
            AssemblyCopyToolForm.AdjustRootDirectory()
        End If

        Me.Close()
    End Sub
End Class