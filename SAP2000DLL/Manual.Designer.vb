<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Manual
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.btnGroup = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnCheckBeamDepth = New System.Windows.Forms.Button()
        Me.btnSplitColumn = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnReactionForce = New System.Windows.Forms.Button()
        Me.btnSidesway = New System.Windows.Forms.Button()
        Me.btnStartEnd = New System.Windows.Forms.Button()
        Me.EventLog1 = New System.Diagnostics.EventLog()
        Me.btnPool = New System.Windows.Forms.Button()
        Me.btnSteelRatio = New System.Windows.Forms.Button()
        Me.btnPRLoad_Poind = New System.Windows.Forms.Button()
        Me.btnPRLoad_Distributed = New System.Windows.Forms.Button()
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.btnWindArea = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnGroup
        '
        Me.btnGroup.Location = New System.Drawing.Point(12, 23)
        Me.btnGroup.Name = "btnGroup"
        Me.btnGroup.Size = New System.Drawing.Size(113, 46)
        Me.btnGroup.TabIndex = 0
        Me.btnGroup.Text = "Group"
        Me.btnGroup.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 75)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(113, 46)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Deflection"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'btnCheckBeamDepth
        '
        Me.btnCheckBeamDepth.Location = New System.Drawing.Point(12, 127)
        Me.btnCheckBeamDepth.Name = "btnCheckBeamDepth"
        Me.btnCheckBeamDepth.Size = New System.Drawing.Size(113, 46)
        Me.btnCheckBeamDepth.TabIndex = 2
        Me.btnCheckBeamDepth.Text = "Check Beam Depth"
        Me.btnCheckBeamDepth.UseVisualStyleBackColor = True
        '
        'btnSplitColumn
        '
        Me.btnSplitColumn.Location = New System.Drawing.Point(12, 179)
        Me.btnSplitColumn.Name = "btnSplitColumn"
        Me.btnSplitColumn.Size = New System.Drawing.Size(113, 46)
        Me.btnSplitColumn.TabIndex = 3
        Me.btnSplitColumn.Text = "Split Column"
        Me.btnSplitColumn.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(318, 301)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(113, 46)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnReactionForce
        '
        Me.btnReactionForce.Location = New System.Drawing.Point(12, 231)
        Me.btnReactionForce.Name = "btnReactionForce"
        Me.btnReactionForce.Size = New System.Drawing.Size(113, 46)
        Me.btnReactionForce.TabIndex = 5
        Me.btnReactionForce.Text = "Reaction Force"
        Me.btnReactionForce.UseVisualStyleBackColor = True
        '
        'btnSidesway
        '
        Me.btnSidesway.Location = New System.Drawing.Point(12, 283)
        Me.btnSidesway.Name = "btnSidesway"
        Me.btnSidesway.Size = New System.Drawing.Size(113, 46)
        Me.btnSidesway.TabIndex = 6
        Me.btnSidesway.Text = "Sidesway Check"
        Me.btnSidesway.UseVisualStyleBackColor = True
        '
        'btnStartEnd
        '
        Me.btnStartEnd.Location = New System.Drawing.Point(141, 23)
        Me.btnStartEnd.Name = "btnStartEnd"
        Me.btnStartEnd.Size = New System.Drawing.Size(113, 46)
        Me.btnStartEnd.TabIndex = 7
        Me.btnStartEnd.Text = "Start End Rule Check"
        Me.btnStartEnd.UseVisualStyleBackColor = True
        '
        'EventLog1
        '
        Me.EventLog1.SynchronizingObject = Me
        '
        'btnPool
        '
        Me.btnPool.Location = New System.Drawing.Point(141, 75)
        Me.btnPool.Name = "btnPool"
        Me.btnPool.Size = New System.Drawing.Size(113, 46)
        Me.btnPool.TabIndex = 8
        Me.btnPool.Text = "Create Pool"
        Me.btnPool.UseVisualStyleBackColor = True
        '
        'btnSteelRatio
        '
        Me.btnSteelRatio.Location = New System.Drawing.Point(141, 127)
        Me.btnSteelRatio.Name = "btnSteelRatio"
        Me.btnSteelRatio.Size = New System.Drawing.Size(113, 46)
        Me.btnSteelRatio.TabIndex = 9
        Me.btnSteelRatio.Text = "Steel Ratio"
        Me.btnSteelRatio.UseVisualStyleBackColor = True
        '
        'btnPRLoad_Poind
        '
        Me.btnPRLoad_Poind.Location = New System.Drawing.Point(141, 179)
        Me.btnPRLoad_Poind.Name = "btnPRLoad_Poind"
        Me.btnPRLoad_Poind.Size = New System.Drawing.Size(113, 46)
        Me.btnPRLoad_Poind.TabIndex = 10
        Me.btnPRLoad_Poind.Text = "Loading Input for Piperack(Point)"
        Me.btnPRLoad_Poind.UseVisualStyleBackColor = True
        '
        'btnPRLoad_Distributed
        '
        Me.btnPRLoad_Distributed.Location = New System.Drawing.Point(141, 231)
        Me.btnPRLoad_Distributed.Name = "btnPRLoad_Distributed"
        Me.btnPRLoad_Distributed.Size = New System.Drawing.Size(113, 46)
        Me.btnPRLoad_Distributed.TabIndex = 11
        Me.btnPRLoad_Distributed.Text = "Loading Input for Piperack(Distributed)"
        Me.btnPRLoad_Distributed.UseVisualStyleBackColor = True
        '
        'btnWindArea
        '
        Me.btnWindArea.Location = New System.Drawing.Point(141, 283)
        Me.btnWindArea.Name = "btnWindArea"
        Me.btnWindArea.Size = New System.Drawing.Size(113, 46)
        Me.btnWindArea.TabIndex = 12
        Me.btnWindArea.Text = "Wind Area Calculation"
        Me.btnWindArea.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(396, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "V20.0"
        '
        'Manual
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(443, 371)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnWindArea)
        Me.Controls.Add(Me.btnPRLoad_Distributed)
        Me.Controls.Add(Me.btnPRLoad_Poind)
        Me.Controls.Add(Me.btnSteelRatio)
        Me.Controls.Add(Me.btnPool)
        Me.Controls.Add(Me.btnStartEnd)
        Me.Controls.Add(Me.btnSidesway)
        Me.Controls.Add(Me.btnReactionForce)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSplitColumn)
        Me.Controls.Add(Me.btnCheckBeamDepth)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btnGroup)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "Manual"
        Me.Text = "User Manual"
        CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnGroup As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents btnCheckBeamDepth As System.Windows.Forms.Button
    Friend WithEvents btnSplitColumn As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnReactionForce As System.Windows.Forms.Button
    Friend WithEvents btnSidesway As System.Windows.Forms.Button
    Friend WithEvents btnStartEnd As System.Windows.Forms.Button
    Friend WithEvents EventLog1 As System.Diagnostics.EventLog
    Friend WithEvents btnPool As System.Windows.Forms.Button
    Friend WithEvents btnSteelRatio As System.Windows.Forms.Button
    Friend WithEvents btnPRLoad_Distributed As System.Windows.Forms.Button
    Friend WithEvents btnPRLoad_Poind As System.Windows.Forms.Button
    Friend WithEvents SerialPort1 As System.IO.Ports.SerialPort
    Friend WithEvents btnWindArea As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
