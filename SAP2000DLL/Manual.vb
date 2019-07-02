Public Class Manual

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Deflection Check.pdf")
    End Sub


    Private Sub btnGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroup.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Group.pdf")
    End Sub

    Private Sub btnCheckBeamDepth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckBeamDepth.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Check Beam Depth.pdf")
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnSplitColumn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSplitColumn.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Split Column.pdf")
    End Sub

    Private Sub btnReactionForce_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReactionForce.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Reaction Force.pdf")
    End Sub

    Private Sub btnSidesway_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSidesway.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Sidesway Check.pdf")
    End Sub

    Private Sub btnStartEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStartEnd.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Start End Rule Check.pdf")
    End Sub

    Private Sub btnPool_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPool.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Create Pool.pdf")
    End Sub

    Private Sub btnSteelRatio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSteelRatio.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Steel Ratio.pdf")
    End Sub

    Private Sub btnPRLoad_Distributed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPRLoad_Distributed.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Loading Input for Piperack(Distributed).pdf")
    End Sub

    Private Sub btnPRLoad_Poind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPRLoad_Poind.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Loading Input for Piperack(Point).pdf")
    End Sub

    Private Sub btnWindArea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWindArea.Click
        Process.Start("\\c1199\util\SAP2000_Plugin\Manual\SAP2000 Plugin-Wind Area Calculation.pdf")
    End Sub


    Private Sub Manual_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.btnPRLoad_Distributed.Text = "Loading Input for" & vbCrLf & "Piperack(Distributed)"
        Me.btnPRLoad_Poind.Text = "Loading Input for" & vbCrLf & "Piperack(Point)"
        Me.btnWindArea.Text = "Wind Area" & vbCrLf & "Calculation"
    End Sub

   
End Class