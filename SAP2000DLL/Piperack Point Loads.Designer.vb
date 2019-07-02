<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Piperack_Point_Loads
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LoadPatterncomb = New System.Windows.Forms.ComboBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Unitcomb = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.RadioButton7 = New System.Windows.Forms.RadioButton
        Me.RadioButton6 = New System.Windows.Forms.RadioButton
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Endlength = New System.Windows.Forms.TextBox
        Me.beginLength = New System.Windows.Forms.TextBox
        Me.Load4 = New System.Windows.Forms.TextBox
        Me.Load3 = New System.Windows.Forms.TextBox
        Me.Load2 = New System.Windows.Forms.TextBox
        Me.Load1 = New System.Windows.Forms.TextBox
        Me.Distance4 = New System.Windows.Forms.TextBox
        Me.Distance3 = New System.Windows.Forms.TextBox
        Me.Distance2 = New System.Windows.Forms.TextBox
        Me.Distance1 = New System.Windows.Forms.TextBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.RadioButton5 = New System.Windows.Forms.RadioButton
        Me.RadioButton4 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LoadPatterncomb)
        Me.GroupBox1.Location = New System.Drawing.Point(11, 15)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(255, 52)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Load Pattern Name"
        '
        'LoadPatterncomb
        '
        Me.LoadPatterncomb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.LoadPatterncomb.FormattingEnabled = True
        Me.LoadPatterncomb.Location = New System.Drawing.Point(23, 21)
        Me.LoadPatterncomb.Name = "LoadPatterncomb"
        Me.LoadPatterncomb.Size = New System.Drawing.Size(149, 20)
        Me.LoadPatterncomb.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Unitcomb)
        Me.GroupBox2.Location = New System.Drawing.Point(272, 15)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(144, 52)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Units"
        '
        'Unitcomb
        '
        Me.Unitcomb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Unitcomb.FormattingEnabled = True
        Me.Unitcomb.Items.AddRange(New Object() {"lb,in,F", "lb,ft,F", "kip,in,F", "kip,ft,F", "kN,mm,C", "kN,m,C", "kgf,mm,C", "kgf,m,C", "N,mm,C", "N,m,C", "Ton,mm,C", "Ton,m,C", "kN,cm,C", "kgf,cm,C", "N,cm,C", "Ton,cm,C"})
        Me.Unitcomb.Location = New System.Drawing.Point(21, 21)
        Me.Unitcomb.Name = "Unitcomb"
        Me.Unitcomb.Size = New System.Drawing.Size(103, 20)
        Me.Unitcomb.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ComboBox3)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.RadioButton2)
        Me.GroupBox3.Controls.Add(Me.RadioButton1)
        Me.GroupBox3.Location = New System.Drawing.Point(11, 73)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(222, 92)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Load Type and Direction"
        '
        'ComboBox3
        '
        Me.ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Items.AddRange(New Object() {"X", "Y", "Z", "Gravity"})
        Me.ComboBox3.Location = New System.Drawing.Point(89, 52)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(108, 20)
        Me.ComboBox3.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 12)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Direction"
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(106, 22)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(66, 16)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Moments"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(23, 22)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(53, 16)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Forces"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.RadioButton7)
        Me.GroupBox4.Controls.Add(Me.RadioButton6)
        Me.GroupBox4.Controls.Add(Me.Label9)
        Me.GroupBox4.Controls.Add(Me.Label8)
        Me.GroupBox4.Controls.Add(Me.Label7)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Endlength)
        Me.GroupBox4.Controls.Add(Me.beginLength)
        Me.GroupBox4.Controls.Add(Me.Load4)
        Me.GroupBox4.Controls.Add(Me.Load3)
        Me.GroupBox4.Controls.Add(Me.Load2)
        Me.GroupBox4.Controls.Add(Me.Load1)
        Me.GroupBox4.Controls.Add(Me.Distance4)
        Me.GroupBox4.Controls.Add(Me.Distance3)
        Me.GroupBox4.Controls.Add(Me.Distance2)
        Me.GroupBox4.Controls.Add(Me.Distance1)
        Me.GroupBox4.Location = New System.Drawing.Point(11, 171)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(405, 171)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Point Loads"
        '
        'RadioButton7
        '
        Me.RadioButton7.AutoSize = True
        Me.RadioButton7.Checked = True
        Me.RadioButton7.Location = New System.Drawing.Point(200, 136)
        Me.RadioButton7.Name = "RadioButton7"
        Me.RadioButton7.Size = New System.Drawing.Size(162, 16)
        Me.RadioButton7.TabIndex = 19
        Me.RadioButton7.TabStop = True
        Me.RadioButton7.Text = "Absolute Distance from End-I"
        Me.RadioButton7.UseVisualStyleBackColor = True
        '
        'RadioButton6
        '
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Location = New System.Drawing.Point(25, 136)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(159, 16)
        Me.RadioButton6.TabIndex = 18
        Me.RadioButton6.Text = "Relative Distance from End-I"
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(347, 18)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(14, 12)
        Me.Label9.TabIndex = 17
        Me.Label9.Text = "4."
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(275, 18)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(14, 12)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "3."
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(210, 18)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(14, 12)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "2."
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(137, 18)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(14, 12)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "1."
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(183, 107)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 12)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "End Length"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 107)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 12)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Begin Length"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 73)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 12)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Load(W/Length)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 12)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Distance"
        '
        'Endlength
        '
        Me.Endlength.Location = New System.Drawing.Point(253, 104)
        Me.Endlength.Name = "Endlength"
        Me.Endlength.Size = New System.Drawing.Size(60, 22)
        Me.Endlength.TabIndex = 9
        Me.Endlength.Text = "0"
        '
        'beginLength
        '
        Me.beginLength.Location = New System.Drawing.Point(112, 104)
        Me.beginLength.Name = "beginLength"
        Me.beginLength.Size = New System.Drawing.Size(60, 22)
        Me.beginLength.TabIndex = 8
        Me.beginLength.Text = "0"
        '
        'Load4
        '
        Me.Load4.Location = New System.Drawing.Point(325, 73)
        Me.Load4.Name = "Load4"
        Me.Load4.Size = New System.Drawing.Size(60, 22)
        Me.Load4.TabIndex = 7
        Me.Load4.Text = "0."
        '
        'Load3
        '
        Me.Load3.Location = New System.Drawing.Point(253, 70)
        Me.Load3.Name = "Load3"
        Me.Load3.Size = New System.Drawing.Size(60, 22)
        Me.Load3.TabIndex = 6
        Me.Load3.Text = "0."
        '
        'Load2
        '
        Me.Load2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Load2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList
        Me.Load2.Location = New System.Drawing.Point(185, 70)
        Me.Load2.Name = "Load2"
        Me.Load2.Size = New System.Drawing.Size(60, 22)
        Me.Load2.TabIndex = 5
        Me.Load2.Text = "0."
        '
        'Load1
        '
        Me.Load1.Location = New System.Drawing.Point(112, 70)
        Me.Load1.Name = "Load1"
        Me.Load1.Size = New System.Drawing.Size(60, 22)
        Me.Load1.TabIndex = 4
        Me.Load1.Text = "0."
        '
        'Distance4
        '
        Me.Distance4.Location = New System.Drawing.Point(325, 36)
        Me.Distance4.Name = "Distance4"
        Me.Distance4.Size = New System.Drawing.Size(60, 22)
        Me.Distance4.TabIndex = 3
        Me.Distance4.Text = "0."
        '
        'Distance3
        '
        Me.Distance3.Location = New System.Drawing.Point(253, 36)
        Me.Distance3.Name = "Distance3"
        Me.Distance3.Size = New System.Drawing.Size(60, 22)
        Me.Distance3.TabIndex = 2
        Me.Distance3.Text = "0."
        '
        'Distance2
        '
        Me.Distance2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Distance2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.RecentlyUsedList
        Me.Distance2.Location = New System.Drawing.Point(185, 36)
        Me.Distance2.Name = "Distance2"
        Me.Distance2.Size = New System.Drawing.Size(60, 22)
        Me.Distance2.TabIndex = 1
        Me.Distance2.Text = "0."
        '
        'Distance1
        '
        Me.Distance1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Distance1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.RecentlyUsedList
        Me.Distance1.Location = New System.Drawing.Point(112, 36)
        Me.Distance1.Name = "Distance1"
        Me.Distance1.Size = New System.Drawing.Size(60, 22)
        Me.Distance1.TabIndex = 0
        Me.Distance1.Text = "0."
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.RadioButton5)
        Me.GroupBox5.Controls.Add(Me.RadioButton4)
        Me.GroupBox5.Controls.Add(Me.RadioButton3)
        Me.GroupBox5.Location = New System.Drawing.Point(239, 73)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(177, 92)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Options"
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(23, 65)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(124, 16)
        Me.RadioButton5.TabIndex = 2
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.Text = "Delete Existing Loads"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(23, 43)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(132, 16)
        Me.RadioButton4.TabIndex = 1
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "Replace Existing Loads"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Checked = True
        Me.RadioButton3.Location = New System.Drawing.Point(23, 21)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(123, 16)
        Me.RadioButton3.TabIndex = 0
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Add to Existing loads"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(202, 348)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(85, 32)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(311, 348)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(85, 32)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Button2)
        Me.GroupBox6.Controls.Add(Me.Button1)
        Me.GroupBox6.Controls.Add(Me.GroupBox5)
        Me.GroupBox6.Controls.Add(Me.GroupBox4)
        Me.GroupBox6.Controls.Add(Me.GroupBox3)
        Me.GroupBox6.Controls.Add(Me.GroupBox2)
        Me.GroupBox6.Controls.Add(Me.GroupBox1)
        Me.GroupBox6.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(430, 406)
        Me.GroupBox6.TabIndex = 7
        Me.GroupBox6.TabStop = False
        '
        'Piperack_Point_Loads
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(450, 431)
        Me.Controls.Add(Me.GroupBox6)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "Piperack_Point_Loads"
        Me.Text = "Regular Point Loads"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LoadPatterncomb As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Unitcomb As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Endlength As System.Windows.Forms.TextBox
    Friend WithEvents beginLength As System.Windows.Forms.TextBox
    Friend WithEvents Load4 As System.Windows.Forms.TextBox
    Friend WithEvents Load3 As System.Windows.Forms.TextBox
    Friend WithEvents Load2 As System.Windows.Forms.TextBox
    Friend WithEvents Load1 As System.Windows.Forms.TextBox
    Friend WithEvents Distance4 As System.Windows.Forms.TextBox
    Friend WithEvents Distance3 As System.Windows.Forms.TextBox
    Friend WithEvents Distance2 As System.Windows.Forms.TextBox
    Friend WithEvents Distance1 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton6 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton7 As System.Windows.Forms.RadioButton
End Class
