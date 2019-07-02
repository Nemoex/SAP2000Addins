<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class K_Calc
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

    '注意:  以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnunbrace = New System.Windows.Forms.RadioButton()
        Me.btnbraced = New System.Windows.Forms.RadioButton()
        Me.btnSelectMain = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtBoxMain = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnSelectGASub = New System.Windows.Forms.Button()
        Me.txtGAele = New System.Windows.Forms.TextBox()
        Me.btnFixed = New System.Windows.Forms.RadioButton()
        Me.btnHinge = New System.Windows.Forms.RadioButton()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnSelectGBSub = New System.Windows.Forms.Button()
        Me.txtGBele = New System.Windows.Forms.TextBox()
        Me.txtCalcK = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCalcGB = New System.Windows.Forms.TextBox()
        Me.txtCalcGA = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.btnMajor = New System.Windows.Forms.RadioButton()
        Me.btnMinor = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnunbrace)
        Me.GroupBox1.Controls.Add(Me.btnbraced)
        Me.GroupBox1.Location = New System.Drawing.Point(45, 31)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(122, 89)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Brace Condition"
        '
        'btnunbrace
        '
        Me.btnunbrace.AutoSize = True
        Me.btnunbrace.Location = New System.Drawing.Point(24, 57)
        Me.btnunbrace.Name = "btnunbrace"
        Me.btnunbrace.Size = New System.Drawing.Size(70, 17)
        Me.btnunbrace.TabIndex = 1
        Me.btnunbrace.TabStop = True
        Me.btnunbrace.Text = "unbraced"
        Me.btnunbrace.UseVisualStyleBackColor = True
        '
        'btnbraced
        '
        Me.btnbraced.AutoSize = True
        Me.btnbraced.Location = New System.Drawing.Point(24, 24)
        Me.btnbraced.Name = "btnbraced"
        Me.btnbraced.Size = New System.Drawing.Size(59, 17)
        Me.btnbraced.TabIndex = 0
        Me.btnbraced.TabStop = True
        Me.btnbraced.Text = "Braced"
        Me.btnbraced.UseVisualStyleBackColor = True
        '
        'btnSelectMain
        '
        Me.btnSelectMain.Location = New System.Drawing.Point(182, 15)
        Me.btnSelectMain.Name = "btnSelectMain"
        Me.btnSelectMain.Size = New System.Drawing.Size(81, 34)
        Me.btnSelectMain.TabIndex = 1
        Me.btnSelectMain.Text = "Select"
        Me.btnSelectMain.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtBoxMain)
        Me.GroupBox2.Controls.Add(Me.btnSelectMain)
        Me.GroupBox2.Location = New System.Drawing.Point(45, 131)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(296, 67)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Main Element"
        '
        'txtBoxMain
        '
        Me.txtBoxMain.Enabled = False
        Me.txtBoxMain.Location = New System.Drawing.Point(15, 23)
        Me.txtBoxMain.Name = "txtBoxMain"
        Me.txtBoxMain.Size = New System.Drawing.Size(130, 20)
        Me.txtBoxMain.TabIndex = 2
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.btnSelectGASub)
        Me.GroupBox3.Controls.Add(Me.txtGAele)
        Me.GroupBox3.Controls.Add(Me.btnFixed)
        Me.GroupBox3.Controls.Add(Me.btnHinge)
        Me.GroupBox3.Location = New System.Drawing.Point(45, 216)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(296, 95)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "GA"
        '
        'btnSelectGASub
        '
        Me.btnSelectGASub.Location = New System.Drawing.Point(209, 33)
        Me.btnSelectGASub.Name = "btnSelectGASub"
        Me.btnSelectGASub.Size = New System.Drawing.Size(81, 34)
        Me.btnSelectGASub.TabIndex = 5
        Me.btnSelectGASub.Text = "Calc GA"
        Me.btnSelectGASub.UseVisualStyleBackColor = True
        '
        'txtGAele
        '
        Me.txtGAele.Location = New System.Drawing.Point(76, 66)
        Me.txtGAele.Name = "txtGAele"
        Me.txtGAele.Size = New System.Drawing.Size(121, 20)
        Me.txtGAele.TabIndex = 2
        '
        'btnFixed
        '
        Me.btnFixed.AutoSize = True
        Me.btnFixed.Location = New System.Drawing.Point(15, 42)
        Me.btnFixed.Name = "btnFixed"
        Me.btnFixed.Size = New System.Drawing.Size(50, 17)
        Me.btnFixed.TabIndex = 1
        Me.btnFixed.TabStop = True
        Me.btnFixed.Text = "Fixed"
        Me.btnFixed.UseVisualStyleBackColor = True
        '
        'btnHinge
        '
        Me.btnHinge.AutoSize = True
        Me.btnHinge.Location = New System.Drawing.Point(15, 19)
        Me.btnHinge.Name = "btnHinge"
        Me.btnHinge.Size = New System.Drawing.Size(53, 17)
        Me.btnHinge.TabIndex = 0
        Me.btnHinge.TabStop = True
        Me.btnHinge.Text = "Hinge"
        Me.btnHinge.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.btnSelectGBSub)
        Me.GroupBox4.Controls.Add(Me.txtGBele)
        Me.GroupBox4.Location = New System.Drawing.Point(45, 317)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(296, 59)
        Me.GroupBox4.TabIndex = 4
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "GB"
        '
        'btnSelectGBSub
        '
        Me.btnSelectGBSub.Location = New System.Drawing.Point(209, 11)
        Me.btnSelectGBSub.Name = "btnSelectGBSub"
        Me.btnSelectGBSub.Size = New System.Drawing.Size(81, 34)
        Me.btnSelectGBSub.TabIndex = 6
        Me.btnSelectGBSub.Text = "Calc GB"
        Me.btnSelectGBSub.UseVisualStyleBackColor = True
        '
        'txtGBele
        '
        Me.txtGBele.Location = New System.Drawing.Point(76, 19)
        Me.txtGBele.Name = "txtGBele"
        Me.txtGBele.Size = New System.Drawing.Size(121, 20)
        Me.txtGBele.TabIndex = 3
        '
        'txtCalcK
        '
        Me.txtCalcK.Location = New System.Drawing.Point(410, 178)
        Me.txtCalcK.Name = "txtCalcK"
        Me.txtCalcK.Size = New System.Drawing.Size(89, 20)
        Me.txtCalcK.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(381, 181)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(23, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "K ="
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(391, 217)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(108, 52)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Calculate"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(373, 117)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(31, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "GA ="
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(373, 146)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(31, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "GB ="
        '
        'txtCalcGB
        '
        Me.txtCalcGB.Location = New System.Drawing.Point(410, 143)
        Me.txtCalcGB.Name = "txtCalcGB"
        Me.txtCalcGB.Size = New System.Drawing.Size(89, 20)
        Me.txtCalcGB.TabIndex = 10
        '
        'txtCalcGA
        '
        Me.txtCalcGA.Location = New System.Drawing.Point(410, 114)
        Me.txtCalcGA.Name = "txtCalcGA"
        Me.txtCalcGA.Size = New System.Drawing.Size(89, 20)
        Me.txtCalcGA.TabIndex = 11
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.btnMinor)
        Me.GroupBox5.Controls.Add(Me.btnMajor)
        Me.GroupBox5.Location = New System.Drawing.Point(198, 31)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(110, 89)
        Me.GroupBox5.TabIndex = 12
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Column Direction"
        '
        'btnMajor
        '
        Me.btnMajor.AutoSize = True
        Me.btnMajor.Location = New System.Drawing.Point(6, 24)
        Me.btnMajor.Name = "btnMajor"
        Me.btnMajor.Size = New System.Drawing.Size(51, 17)
        Me.btnMajor.TabIndex = 0
        Me.btnMajor.TabStop = True
        Me.btnMajor.Text = "Major"
        Me.btnMajor.UseVisualStyleBackColor = True
        '
        'btnMinor
        '
        Me.btnMinor.AutoSize = True
        Me.btnMinor.Location = New System.Drawing.Point(6, 57)
        Me.btnMinor.Name = "btnMinor"
        Me.btnMinor.Size = New System.Drawing.Size(51, 17)
        Me.btnMinor.TabIndex = 1
        Me.btnMinor.TabStop = True
        Me.btnMinor.Text = "Minor"
        Me.btnMinor.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(21, 69)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Selected"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(19, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Selected"
        '
        'K_Calc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(595, 388)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.txtCalcGA)
        Me.Controls.Add(Me.txtCalcGB)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCalcK)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "K_Calc"
        Me.Text = "K_Calc"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnunbrace As System.Windows.Forms.RadioButton
    Friend WithEvents btnbraced As System.Windows.Forms.RadioButton
    Friend WithEvents btnSelectMain As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtBoxMain As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSelectGASub As System.Windows.Forms.Button
    Friend WithEvents txtGAele As System.Windows.Forms.TextBox
    Friend WithEvents btnFixed As System.Windows.Forms.RadioButton
    Friend WithEvents btnHinge As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSelectGBSub As System.Windows.Forms.Button
    Friend WithEvents txtGBele As System.Windows.Forms.TextBox
    Friend WithEvents txtCalcK As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtCalcGB As System.Windows.Forms.TextBox
    Friend WithEvents txtCalcGA As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btnMinor As System.Windows.Forms.RadioButton
    Friend WithEvents btnMajor As System.Windows.Forms.RadioButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
