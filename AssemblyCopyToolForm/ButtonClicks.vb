Imports System.IO
Module ButtonClicks
    Sub CopyButtonHandler(frm As AssemblyCopyToolForm, rootAssemblyObject As AssemblyCopyObject, invApp As Inventor.Application,
                          sender As Object, e As EventArgs)
        AssemblyCopyToolForm.LB_CopyComplete.Visible = True
        frm.LB_CopyComplete.Text = "Starting Process..."
        rootAssemblyObject.UpdateNewProperties()
        rootAssemblyObject.CreateNewFiles(dryrun:=False)

        'try closing the original assembly without saving to see if this fixes the platform replacement issue.
        invApp.ActiveDocument.Close(False)

        rootAssemblyObject.ReplaceOccurrencesByIndex()
        frm.LB_CopyComplete.Text = "Assembly Copy Complete!"
        invApp.ActiveDocument.Save2()
    End Sub

    Sub NewDirectoryButtonHandler(frm As AssemblyCopyToolForm, sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select New Root Directory for Copied Assembly"
            fbd.ShowNewFolderButton = True
            If fbd.ShowDialog() = DialogResult.OK Then
                frm.TB_newDir.Text = fbd.SelectedPath
            End If
        End Using
    End Sub

    Sub BT_PrefixSuffixHandler(frm As AssemblyCopyToolForm, sender As Object, e As EventArgs)
        Dim node As TreeNode = frm.TV_nComponent.SelectedNode
        Dim oNodeText As String = node.Text
        node.Text = frm.TB_Prefix.Text & oNodeText & frm.TB_Suffix.Text
    End Sub

    Sub TestButtonClickHandler(sender As Object, e As EventArgs, Optional frm As AssemblyCopyToolForm = Nothing, Optional rootAssemblyObject As AssemblyCopyObject = Nothing, Optional invApp As Inventor.Application = Nothing)
        'NameFirstOccurrence(invApp)
        ReadSelectionAttributes(invApp)

    End Sub


End Module
