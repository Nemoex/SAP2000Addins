Public Class MemberEndForce

   
    Private Sub MemberEndForce_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SaveFileDialog1.FileName = "DeflectionCheck.txt"

        SaveFileDialog1.ShowDialog()

        TextBox1.Text = SaveFileDialog1.FileName
    End Sub

   
End Class