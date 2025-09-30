Module CopyToolFunctions
    Sub SaveAssemblyFile(ByRef invtAsmObj As InvtAssemblyObj)
        DB("Saving Assembly: " & invtAsmObj.NewFileName)
        Dim rootDirectory As String = invtAsmObj.NewFilePath
        DB(invtAsmObj.NewFullFileName)
        DB(" ")


        'For now all new components will have the same name as the original
        For Each comp As InvtComponentObj In invtAsmObj.AssemblyComponents
            If comp.Type = "Part" AndAlso comp.PartObject IsNot Nothing Then
                Dim curPrtObj As InvtPartObj = comp.PartObject
                curPrtObj.NewName = comp.PartObject.OriginalName
                curPrtObj.NewFilePath = rootDirectory
                DB("Saving Part: " & curPrtObj.NewFileName)
                DB(curPrtObj.NewFullFileName)
                DB("")
            ElseIf comp.Type = "Assembly" AndAlso comp.AssemblyObject IsNot Nothing Then
                Dim curAsmObj As InvtAssemblyObj = comp.AssemblyObject
                curAsmObj.NewName = curAsmObj.OriginalName
                curAsmObj.NewFilePath = rootDirectory
                SaveAssemblyFile(curAsmObj)

            End If
        Next

    End Sub

    ''' <summary>
    ''' Writes text in a textbox and scrolls to the end
    ''' </summary>
    ''' <param name="textBox"></param>
    ''' <param name="msg"></param>
    Sub LongTextboxWrite(ByRef textBox As System.Windows.Forms.TextBox, ByRef msg As String)
        textBox.Text = msg
        textBox.SelectionStart = textBox.Text.Length
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
    End Sub

    Sub MoveCaret(ByRef textBox As System.Windows.Forms.TextBox)
        textBox.SelectionStart = 0
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
        textBox.SelectionStart = textBox.Text.Length
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
    End Sub


End Module
