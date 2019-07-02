Public Class SelectCombDialog



    Private Sub SelectCombDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TextBox1.Text = My.Computer.FileSystem.CurrentDirectory + "\DeflectionCheck.txt"

        SaveFileDialog1.FileName = "DeflectionCheck.txt"

        SaveFileDialog1.ShowDialog()

        TextBox1.Text = SaveFileDialog1.FileName

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        ReportFileName = TextBox1.Text

        If DefChkCriteria = True Then
            DeflectionCriteria = TextBox2.Text
        Else
            DeflectionCriteria = TextBox4.Text
        End If

        ReDim CombList(ListBox2.Items.Count - 1)

        For i = 0 To ListBox2.Items.Count - 1
            CombList(i) = ListBox2.Items(i)
        Next

        If ListBox2.Items.Count = 0 Then
            MsgBox("No Select any Combination", MsgBoxStyle.Question, "Deflection Check")
            SelectCombFlag = False
            GoTo ExitSub
        Else
            SelectCombFlag = True
        End If

        If IsNumeric(TextBox2.Text) = False Then
            MsgBox("Limit is not a Numeric!")
            Exit Sub
        End If

        Try
            FileOpen(20, ReportFileName, OpenMode.Output)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Deflection Check")
            FileClose()
            OpenFileErrorFlag = True
            Exit Sub
        End Try

        If TextBox3.Text.Contains("Double Click") = False And TextBox3.Text <> "" Then
            JointSequenceFile = TextBox3.Text
        Else
            JointSequenceFile = ""
        End If



        If JointSequenceFile <> "" Then
            Try
                FileOpen(40, JointSequenceFile, OpenMode.Input)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Deflection Check")

                FileClose()
                OpenFileErrorFlag = True
                GoTo ExitSub
            End Try
        Else
            MsgBox("There maybe some inaccuracy at AutoMesh Point because you didn't load ~JointTable" & vbCrLf & "Set Unit to mm" & vbCrLf & "Display > Show Tables" & vbCrLf & "ANALYSIS RESULTS" & vbCrLf & " >Element Output" & vbCrLf & "  >Objects and Elements" & vbCrLf & "   > Table : Object And Elements - Joints", MsgBoxStyle.Information + MsgBoxStyle.SystemModal, "Deflection Check Notice")
        End If




ExitSub:
        FileClose(20)
        FileClose(40)
        'My.Computer.FileSystem.DeleteFile(ReportFileName)

        Me.Close()

    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
      
        While ListBox1.SelectedIndices.Count > 0
            ListBox2.Items.Add(ListBox1.Items(ListBox1.SelectedIndices(0)))
            ListBox1.Items.Remove(ListBox1.Items(ListBox1.SelectedIndices(0)))
        End While

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
       
        While ListBox2.SelectedIndices.Count > 0
            ListBox1.Items.Add(ListBox2.Items(ListBox2.SelectedIndices(0)))
            ListBox2.Items.Remove(ListBox2.Items(ListBox2.SelectedIndices(0)))
        End While

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub


    Private Sub TextBox3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox3.MouseDoubleClick

        OpenFileDialog1.ShowDialog()
        TextBox3.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged

        If RadioButton2.Checked Then
            TextBox2.Enabled = False
            TextBox4.Enabled = True
            DefChkCriteria = False
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            TextBox2.Enabled = True
            TextBox4.Enabled = False
            DefChkCriteria = True
        End If

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub chkShowSectName_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowSectName.CheckedChanged
        If chkShowSectName.Checked Then
            ShowSectName = True
        Else
            ShowSectName = False
        End If
    End Sub
  
    Private Sub chkShowCriNode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowCriNode.CheckedChanged
        If chkShowCriNode.Checked Then
            ShowCriNode = True
        Else
            ShowCriNode = False
        End If
    End Sub

    Private Sub chkShowNodeList_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowNodeList.CheckedChanged
        If chkShowNodeList.Checked Then
            ShowNodeList = True
        Else
            ShowNodeList = False
        End If
    End Sub

  
    Private Sub chkRelativebtn_CheckedChanged(sender As Object, e As EventArgs) Handles chkRelativebtn.CheckedChanged
        If chkRelativebtn.Checked Then
            CheckDeflectionRelative = True
            CheckDeflectionAbsolute = False
        Else
            CheckDeflectionRelative = False
            CheckDeflectionAbsolute = True
        End If
    End Sub

    '不需要重複檢查
    'Private Sub chkAbsolutebtn_CheckedChanged(sender As Object, e As EventArgs) Handles chkAbsolutebtn.CheckedChanged
    '    If chkAbsolutebtn.Checked Then
    '        CheckDeflectionRelative = False
    '        CheckDeflectionAbsolute = True
    '    Else
    '        CheckDeflectionRelative = True
    '        CheckDeflectionAbsolute = False
    '    End If
    'End Sub
End Class