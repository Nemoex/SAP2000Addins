Imports SAP2000v20


Public Class PoolDataForm

    Private Sub PoolDataForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Pool_Length = CDbl(TextBox1.Text)
        Pool_Width = CDbl(TextBox2.Text)
        Pool_Height = CDbl(TextBox3.Text)
        Divide_X = CInt(TextBox4.Text)
        Divide_Y = CInt(TextBox5.Text)
        Divide_Z = CInt(TextBox6.Text)
        WallThk = CDbl(TextBox7.Text)
        BottomSlabThk = CDbl(TextBox9.Text)
        WithTopSlab = CheckBox1.Checked
        TopSlabThk = CDbl(TextBox8.Text)
        TopWallThk = CDbl(TextBox10.Text)
        BottomWallThk = CDbl(TextBox11.Text)
        WaterPressure_Top = CDbl(TextBox12.Text)
        WaterPressure_Bottom = CDbl(TextBox13.Text)
        WaterPressure2_Top = CDbl(TextBox14.Text)
        WaterPressure2_Bottom = CDbl(TextBox15.Text)
        VerSoil_K = CDbl(TextBox16.Text)
        HoriSoil_K = CDbl(TextBox17.Text)
        WaterPressure3_Top = CDbl(TextBox18.Text)
        WaterPressure3_Bottom = CDbl(TextBox19.Text)
        WaterPressure4_Top = CDbl(TextBox20.Text)
        WaterPressure4_Bottom = CDbl(TextBox21.Text)
        WaterDepth = CDbl(TextBox22.Text)
        VerSeismicCoeff = CDbl(TextBox23.Text)

        Me.DialogResult = Windows.Forms.DialogResult.OK

        Me.Close()
    End Sub

    
    
    Private Sub TextBox22_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox22.TextChanged
        If IsNumeric(TextBox22.Text) = False Then
            Button1.Enabled = False
            GoTo 100
        End If

        If TextBox22.Text.Contains("-") Or TextBox22.Text.Contains(" ") Then
            'MsgBox("Please Input Positive Value")
            Button1.Enabled = False
            GoTo 100
        ElseIf TextBox22.Text = "" Then
            Button1.Enabled = False
            GoTo 100
        End If


        If CDbl(TextBox22.Text) > CDbl(TextBox3.Text) Then
            MsgBox("Water depth can't greater then pool height")
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If

100:

    End Sub
End Class