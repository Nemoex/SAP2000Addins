Public Class Select_Rebar_Size

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Dim RebarSize_Girder As String
        'Dim RebarSize_Beam As String
        'Dim RebarSize_Torsion As String

        'Dim RebarArea_Pri As Double
        'Dim RebarArea_Sec As Double
        'Dim RebarArea_Tor As Double
        'Dim RebarArea As Double

        'RebarSize_Girder = InputBox("Rebar Size - Girder " & vbCrLf & "EX : D25 / #8  etc...", "Input Rebar Size - Girder", "D25").ToUpper
        RebarSize_Girder = ComboBox1.SelectedItem

        'RebarSize_Beam = InputBox("Rebar Size - Beam  " & vbCrLf & "EX : D25 / #8 etc...", "Input Rebar Size - Beam", "D25").ToUpper
        'RebarSize_Torsion = InputBox("Rebar Size - Torsion  " & vbCrLf & "EX : D19 / #6  etc...", "Input Rebar Size - Torsion", "D19").ToUpper
        RebarSize_Beam = ComboBox2.SelectedItem
        RebarSize_Torsion = ComboBox3.SelectedItem

        RebarArea_Pri = GetRebarArea(RebarSize_Girder)
        If RebarArea_Pri = -1 Then MsgBox("Can't find " & RebarSize_Girder & " please contact RD Team")
        RebarArea_Sec = GetRebarArea(RebarSize_Beam)
        If RebarArea_Sec = -1 Then MsgBox("Can't find " & RebarSize_Beam & " please contact RD Team")
        RebarArea_Tor = GetRebarArea(RebarSize_Torsion)
        If RebarArea_Tor = -1 Then MsgBox("Can't find " & RebarSize_Torsion & " please contact RD Team")

        If RebarArea_Pri = -1 Or RebarArea_Sec = -1 Or RebarArea_Tor = -1 Then
            GoTo EndSub
        End If



endsub:
        Me.Close()
    End Sub

    
End Class