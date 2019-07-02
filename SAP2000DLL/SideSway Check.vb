Imports System.Windows.Forms
Imports System.Xml

Public Class SideSway_Check

    Public SapModel As SAP2000v20.cSapModel
    Public ISapPlugin As SAP2000v20.cPluginCallback

    Private Sub SideSway_Check_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        PictureBox1.BringToFront()

        Dim ret As Long
        Dim NumberCombs As Integer
        Dim CombName() As String
        ret = SapModel.RespCombo.GetNameList(NumberCombs, CombName)
        For i = 0 To NumberCombs - 1
            ListBox1.Items.Add(CombName(i))
        Next

        Me.Button3.Text = "Paste from" & vbCrLf & "Clipboard"
        Me.Button4.Text = "Clear Grid Line" & vbCrLf & "Data"
        Me.Button8.Text = "Output to Text" & vbCrLf & "File"

    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If DataGridView1.Item(0, 0).Value = "" Then
            MsgBox("Grid Data Is Empty")
            GoTo ExitSub
        End If

        If DataGridView1.Item(2, 1).Value > 500 Then
            MsgBox("Please check unit", MsgBoxStyle.Critical)
        End If


        If ListBox1.Items.Count + ListBox2.Items.Count = 0 Then
            MsgBox("Can't find any load combination", MsgBoxStyle.SystemModal, "Sidesway Check")
            GoTo ExitSub
        End If


        Dim ret As Long

        Dim NumberCombs As Integer

        Dim NumberItems As Integer
        Dim CombName() As String

        Dim CaseName() As String
        Dim Status() As Integer
        Dim AnalyzedFlag As Boolean = False

        ret = SapModel.Analyze.GetCaseStatus(NumberItems, CaseName, Status)

        For i = 0 To NumberItems - 1
            If Status(i) = 4 Then
                AnalyzedFlag = True
            End If
        Next

        If AnalyzedFlag <> True Then
            MsgBox("Need Run Analysis First" & vbCrLf & "End Program", , "Sidesway Check")
            GoTo ExitSub
        End If

        'get all combination 
        ret = SapModel.RespCombo.GetNameList(NumberCombs, CombName)
        ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput

        'get begining units
        Dim beginUnits As SAP2000v20.eUnits
        beginUnits = SapModel.GetPresentUnits()


        'get point displacements
        SapModel.SetPresentUnits(SAP2000v20.eUnits.Ton_m_C)

        Dim VoidPointCounter As Integer

        Dim AutoMesh As Boolean
        Dim AutoMeshAtPoints As Boolean
        Dim AutoMeshAtLines As Boolean
        Dim NumSegs As Long
        Dim AutoMeshMaxLength As Double
        Dim AutoSelectFrame As MsgBoxResult
        Dim ParallelTo(5) As Boolean

        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        If NumberItems = 0 Then
            AutoSelectFrame = MsgBox("Didn't select any frame for check" & vbCrLf & "Do you want to select ""ALL Vertical Frame"" to check ? ", MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal, "Sidesway Check")
            If AutoSelectFrame = MsgBoxResult.Yes Then
                ParallelTo(2) = True
                ret = SapModel.SelectObj.LinesParallelToCoordAxis(ParallelTo)
            Else
                MsgBox("Please select frame first then re-start command", , "Sidesway Check")
                GoTo ExitSub
            End If
        End If

        'define sidesway limit
        'Dim Limitation As String
        'Limitation = InputBox("Input Sidesway limit", "Sidesway Check", "200")

        If IsNumeric(TextBox1.Text) = False Then
            MsgBox("Please check sidesway limit value", MsgBoxStyle.Critical, "Sidesway Check")
            GoTo ExitSub
        Else
            DataGridView2.Columns(2).HeaderText = "H/" + TextBox1.Text + "  (mm)"
        End If


        Dim Xcount As Integer = 0
        Dim Ycount As Integer = 0

        Dim Xtext() As String
        Dim Ytext() As String

        Dim Xpos() As Double
        Dim Ypos() As Double

        Dim DataGridrowCount As Integer = 0

        For i = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(1, i).Value <> "" Then
                DataGridrowCount += 1
            End If

        Next



        For i = 0 To DataGridrowCount - 1
            If DataGridView1.Item(0, i).Value.ToString.ToUpper = "X" Then
                ReDim Preserve Xtext(Xcount)
                ReDim Preserve Xpos(Xcount)
                Xtext(Xcount) = DataGridView1.Item(1, i).Value
                Xpos(Xcount) = DataGridView1.Item(2, i).Value
                Xcount += 1
            End If
            If DataGridView1.Item(0, i).Value.ToString.ToUpper = "Y" Then
                ReDim Preserve Ytext(Ycount)
                ReDim Preserve Ypos(Ycount)
                Ytext(Ycount) = DataGridView1.Item(1, i).Value
                Ypos(Ycount) = DataGridView1.Item(2, i).Value
                Ycount += 1
            End If
        Next

        Dim ColLineCollection(Xcount * Ycount - 1) As GridIntersection

        Dim collcount As Integer = 0

        For i = 0 To Xcount - 1
            For j = 0 To Ycount - 1
                ColLineCollection(collcount).Name = Xtext(i) + " / " + Ytext(j)
                ColLineCollection(collcount).location.X = Xpos(i)
                ColLineCollection(collcount).location.Y = Ypos(j)
                collcount += 1
            Next
        Next

        Dim P1, P2 As String
        Dim X1, X2, Y1, Y2, Z1, Z2 As Double
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)





        For i = 0 To NumberItems - 1
            If ObjectType(i) = 2 Then
                ret = SapModel.FrameObj.GetPoints(ObjectName(i), P1, P2)
                ret = SapModel.PointObj.GetCoordCartesian(P1, X1, Y1, Z1)
                ret = SapModel.PointObj.GetCoordCartesian(P2, X2, Y2, Z2)
                If Math.Abs(X1 - X2) < 0.0001 And Math.Abs(Y1 - Y2) < 0.0001 Then
                    For Each P As GridIntersection In ColLineCollection
                        If Math.Abs(X1 - P.location.X) < 0.0001 And Math.Abs(Y1 - P.location.Y) < 0.0001 Then
                            ReDim Preserve P.Member(P.MemberCount)
                            P.MemberCount += 1
                            P.Member(P.MemberCount - 1) = ObjectName(i)
                            ReDim Preserve P.Joints(P.JointCount)
                            If Array.IndexOf(P.Joints, P1) = -1 Then
                                P.JointCount += 1
                                P.Joints(P.JointCount - 1) = P1
                            End If
                            ReDim Preserve P.Joints(P.JointCount)
                            If Array.IndexOf(P.Joints, P2) = -1 Then

                                P.JointCount += 1
                                P.Joints(P.JointCount - 1) = P2
                            End If

                            For j = 0 To ColLineCollection.Count - 1
                                If P.Name = ColLineCollection(j).Name Then
                                    ColLineCollection(j) = P
                                    Exit For
                                End If
                            Next
                            Exit For
                        End If
                    Next
                End If
            End If
        Next

        Dim NumberResults As Integer
        Dim Obj() As String
        Dim Elm() As String
        Dim LoadCase() As String
        Dim StepType() As String
        Dim StepNum() As Double
        Dim U1() As Double
        Dim U2() As Double
        Dim U3() As Double
        Dim R1() As Double
        Dim R2() As Double
        Dim R3() As Double

        Dim TX1 As Double
        Dim TX2 As Double
        Dim DX As Double
        Dim TY1 As Double
        Dim TY2 As Double
        Dim DY As Double

        Dim DXMAX As Double
        Dim DYMAX As Double
        Dim CtrlCombX As String
        Dim CtrlCombY As String
        Dim TX1MAX, TX2MAX As Double
        Dim TY1MAX, TY2MAX As Double
        Dim ColumnLine As String

        Dim DXYMAX As Double
        Dim CtrlCombXY As String
        Dim DXY As Double



        DataGridView2.Rows.Clear()

        SapModel.SetPresentUnits(SAP2000v20.eUnits.Ton_mm_C)

        If btnMaxLC.Checked Then
            For i = 0 To ColLineCollection.Count - 1
                For l = 0 To ColLineCollection(i).JointCount - 1
                    For m = 0 To ColLineCollection(i).JointCount - 1
                        If l < m Then
                            For n = 0 To ListBox2.Items.Count - 1
                                ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
                                ret = SapModel.Results.Setup.SetComboSelectedForOutput(ListBox2.Items(n))

                                ret = SapModel.Results.JointDispl(ColLineCollection(i).Joints(l), SAP2000v20.eItemTypeElm.Element, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
                                TX1 = U1(0)
                                TY1 = U2(0)
                                ret = SapModel.Results.JointDispl(ColLineCollection(i).Joints(m), SAP2000v20.eItemTypeElm.Element, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
                                TX2 = U1(0)
                                TY2 = U2(0)
                                DX = Math.Abs(TX1 - TX2)
                                DY = Math.Abs(TY1 - TY2)
                                If DXMAX < DX Then
                                    DXMAX = DX
                                    CtrlCombX = ListBox2.Items(n)
                                    TX1MAX = TX1
                                    TX2MAX = TX2
                                End If
                                If DYMAX < DY Then
                                    DYMAX = DY
                                    CtrlCombY = ListBox2.Items(n)
                                    TY1MAX = TY1
                                    TY2MAX = TY2
                                End If

                                DXY = (DX ^ 2 + DY ^ 2) ^ 0.5
                                If DXYMAX < DXY Then
                                    DXYMAX = DXY
                                    CtrlCombXY = ListBox2.Items(n)
                                End If


                            Next
                            '========================此處開始記錄或輸出2點之計算值

                            ret = SapModel.PointObj.GetCoordCartesian(ColLineCollection(i).Joints(l), X1, Y1, Z1)
                            ret = SapModel.PointObj.GetCoordCartesian(ColLineCollection(i).Joints(m), X2, Y2, Z2)

                            Dim dist As Double
                            Dim Pass1, Pass2 As String
                            Dim Pass3 As String
                            dist = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5 / CDbl(TextBox1.Text)
                            If DXMAX < dist Then
                                Pass1 = "OK"
                            Else
                                Pass1 = "NG"
                            End If
                            If DYMAX < dist Then
                                Pass2 = "OK"
                            Else
                                Pass2 = "NG"
                            End If
                            If DXYMAX < dist Then
                                Pass3 = "OK"
                            Else
                                Pass3 = "NG"
                            End If




                            ColumnLine = ColLineCollection(i).Name

                            Dim row As String()
                            row = New String() {ColumnLine, ColLineCollection(i).Joints(l) + " - " + ColLineCollection(i).Joints(m), Format(dist, "F2"), CtrlCombX, Format(TX1MAX, "F2"), Format(TX2MAX, "F2"), Format(DXMAX, "F2"), Pass1, CtrlCombY, Format(TY1MAX, "F2"), Format(TY2MAX, "F2"), Format(DYMAX, "F2"), Pass2, CtrlCombXY, Format(DXYMAX, "F2"), Pass3}
                            DataGridView2.Rows.Add(row)


                            '========================
                            DXMAX = 0
                            DYMAX = 0
                            DXYMAX = 0
                            TX1MAX = 0
                            TX2MAX = 0
                            TY1MAX = 0
                            TY2MAX = 0
                            CtrlCombX = ""
                            CtrlCombY = ""
                            CtrlCombXY = ""
                        End If
                    Next
                Next
            Next


        ElseIf btnALL.Checked Then
            For i = 0 To ColLineCollection.Count - 1
                For l = 0 To ColLineCollection(i).JointCount - 1
                    For m = 0 To ColLineCollection(i).JointCount - 1
                        If l < m Then
                            For n = 0 To ListBox2.Items.Count - 1
                                ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
                                ret = SapModel.Results.Setup.SetComboSelectedForOutput(ListBox2.Items(n))

                                ret = SapModel.Results.JointDispl(ColLineCollection(i).Joints(l), SAP2000v20.eItemTypeElm.Element, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
                                TX1 = U1(0)
                                TY1 = U2(0)
                                ret = SapModel.Results.JointDispl(ColLineCollection(i).Joints(m), SAP2000v20.eItemTypeElm.Element, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
                                TX2 = U1(0)
                                TY2 = U2(0)
                                DX = Math.Abs(TX1 - TX2)
                                DY = Math.Abs(TY1 - TY2)

                                'If DXMAX < DX Then
                                DXMAX = DX
                                CtrlCombX = ListBox2.Items(n)
                                TX1MAX = TX1
                                TX2MAX = TX2
                                'End If
                                'If DYMAX < DY Then
                                DYMAX = DY
                                CtrlCombY = ListBox2.Items(n)
                                TY1MAX = TY1
                                TY2MAX = TY2
                                'End If
                                'If DXYMAX < DXY Then

                                DXY = (DX ^ 2 + DY ^ 2) ^ 0.5

                                DXYMAX = DXY
                                CtrlCombY = ListBox2.Items(n)
                                'End If

                                '==========直接輸出

                                ret = SapModel.PointObj.GetCoordCartesian(ColLineCollection(i).Joints(l), X1, Y1, Z1)
                                ret = SapModel.PointObj.GetCoordCartesian(ColLineCollection(i).Joints(m), X2, Y2, Z2)

                                Dim dist As Double
                                Dim Pass1, Pass2, Pass3 As String
                                dist = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5 / CDbl(TextBox1.Text)
                                If DXMAX < dist Then
                                    Pass1 = "OK"
                                Else
                                    Pass1 = "NG"
                                End If
                                If DYMAX < dist Then
                                    Pass2 = "OK"
                                Else
                                    Pass2 = "NG"
                                End If
                                If DXYMAX < dist Then
                                    Pass3 = "OK"
                                Else
                                    Pass3 = "NG"
                                End If


                                ColumnLine = ColLineCollection(i).Name

                                Dim row As String()
                                row = New String() {ColumnLine, ColLineCollection(i).Joints(l) + " - " + ColLineCollection(i).Joints(m), Format(dist, "F2"), CtrlCombX, Format(TX1MAX, "F2"), Format(TX2MAX, "F2"), Format(DXMAX, "F2"), Pass1, CtrlCombY, Format(TY1MAX, "F2"), Format(TY2MAX, "F2"), Format(DYMAX, "F2"), Pass2, CtrlCombXY, Format(DXYMAX, "F2"), Pass3}
                                DataGridView2.Rows.Add(row)


                                '========================
                                DXMAX = 0
                                DYMAX = 0
                                DXYMAX = 0
                                TX1MAX = 0
                                TX2MAX = 0
                                TY1MAX = 0
                                TY2MAX = 0
                                CtrlCombX = ""
                                CtrlCombY = ""
                                CtrlCombXY = ""


                            Next
                            '========================此處開始記錄或輸出2點之計算值

                            'ret = SapModel.PointObj.GetCoordCartesian(ColLineCollection(i).Joints(l), X1, Y1, Z1)
                            'ret = SapModel.PointObj.GetCoordCartesian(ColLineCollection(i).Joints(m), X2, Y2, Z2)

                            'Dim dist As Double
                            'Dim Pass1, Pass2 As String
                            'dist = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5 / CDbl(TextBox1.Text)
                            'If DXMAX < dist Then
                            '    Pass1 = "OK"
                            'Else
                            '    Pass1 = "NG"
                            'End If
                            'If DYMAX < dist Then
                            '    Pass2 = "OK"
                            'Else
                            '    Pass2 = "NG"
                            'End If

                            'ColumnLine = ColLineCollection(i).Name

                            'Dim row As String()
                            'row = New String() {ColumnLine, ColLineCollection(i).Joints(l) + " - " + ColLineCollection(i).Joints(m), Format(dist, "F2"), CtrlCombX, Format(TX1MAX, "F2"), Format(TX2MAX, "F2"), Format(DXMAX, "F2"), Pass1, CtrlCombY, Format(TY1MAX, "F2"), Format(TY2MAX, "F2"), Format(DYMAX, "F2"), Pass2}
                            'DataGridView2.Rows.Add(row)


                            ''========================
                            'DXMAX = 0
                            'DYMAX = 0
                            'TX1MAX = 0
                            'TX2MAX = 0
                            'TY1MAX = 0
                            'TY2MAX = 0
                            'CtrlCombX = ""
                            'CtrlCombY = ""
                        End If
                    Next
                Next
            Next

        End If


        Me.TabControl1.SelectedTab = TabPage2

ExitSub:
        SapModel.SetPresentUnits(beginUnits)

    End Sub

    Private Sub PasteUnboundRecords()

        Try
            Dim rowLines As String() = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Text).Split(New String(0) {vbCr & vbLf}, StringSplitOptions.None)
            Dim currentRowIndex As Integer = (If(DataGridView1.CurrentRow IsNot Nothing, DataGridView1.CurrentRow.Index, 0))
            Dim currentColumnIndex As Integer = (If(DataGridView1.CurrentCell IsNot Nothing, DataGridView1.CurrentCell.ColumnIndex, 0))
            Dim currentColumnCount As Integer = DataGridView1.Columns.Count

            For i = 1 To rowLines.Count
                DataGridView1.Rows.Add()
            Next

            DataGridView1.AllowUserToAddRows = False
            For rowLine As Integer = 0 To rowLines.Length - 1

                If rowLine = rowLines.Length - 1 AndAlso String.IsNullOrEmpty(rowLines(rowLine)) Then
                    Exit For
                End If

                Dim columnsData As String() = rowLines(rowLine).Split(New String(0) {vbTab}, StringSplitOptions.None)
                If (currentColumnIndex + columnsData.Length) > DataGridView1.Columns.Count Then
                    For columnCreationCounter As Integer = 0 To ((currentColumnIndex + columnsData.Length) - currentColumnCount) - 1
                        If columnCreationCounter = rowLines.Length - 1 Then
                            Exit For
                        End If
                    Next
                End If
                If DataGridView1.Rows.Count > (currentRowIndex + rowLine) Then
                    For columnsDataIndex As Integer = 0 To columnsData.Length - 1
                        If currentColumnIndex + columnsDataIndex <= DataGridView1.Columns.Count - 1 Then
                            DataGridView1.Rows(currentRowIndex + rowLine).Cells(currentColumnIndex + columnsDataIndex).Value = columnsData(columnsDataIndex)
                        End If
                    Next
                Else
                    Dim pasteCells As String() = New String(DataGridView1.Columns.Count - 1) {}
                    For cellStartCounter As Integer = currentColumnIndex To DataGridView1.Columns.Count - 1
                        If columnsData.Length > (cellStartCounter - currentColumnIndex) Then
                            pasteCells(cellStartCounter) = columnsData(cellStartCounter - currentColumnIndex)
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
            'Log Exception
        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub DataGridView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        Try
            If e.Control And (e.KeyCode = Keys.C) Then
                Dim d As System.Windows.Forms.DataObject = DataGridView1.GetClipboardContent()
                System.Windows.Forms.Clipboard.SetDataObject(d)
                e.Handled = True
            ElseIf (e.Control And e.KeyCode = Keys.V) Then
                PasteUnboundRecords()
            End If
        Catch ex As Exception
            'Log Exception
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        PasteUnboundRecords()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        DataGridView1.Rows.Add()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub Label2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseHover
        Dim P As System.Drawing.Point
        P.X = 200
        P.Y = 30
        PictureBox1.Location = P
    End Sub
    Private Sub Label2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseLeave
        Dim P As System.Drawing.Point
        P.X = 2000
        P.Y = 30
        PictureBox1.Location = P
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        While ListBox2.SelectedIndices.Count > 0
            ListBox1.Items.Add(ListBox2.Items(ListBox2.SelectedIndices(0)))
            ListBox2.Items.Remove(ListBox2.Items(ListBox2.SelectedIndices(0)))
        End While
        If ListBox2.Items.Count > 0 Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        While ListBox1.SelectedIndices.Count > 0
            ListBox2.Items.Add(ListBox1.Items(ListBox1.SelectedIndices(0)))
            ListBox1.Items.Remove(ListBox1.Items(ListBox1.SelectedIndices(0)))
        End While
        If ListBox2.Items.Count > 0 Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If
    End Sub

    Public Structure GridIntersection
        Public Name As String
        Public location As Drawing.PointF
        Public MemberCount As Integer
        Public Member() As String
        Public JointCount As Integer
        Public Joints() As String
    End Structure



    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '        If DataGridView2.RowCount <= 1 Then GoTo endsub2

        '        On Error Resume Next
        '        xlApp = CreateObject("Excel.Application")
        '        xlBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value)
        '        '停用警告訊息        
        '        'xlApp.DisplayAlerts = False
        '        '設置EXCEL對象可見        
        '        xlApp.Visible = True
        '        '設定活頁簿為焦點        
        '        xlBook.Activate()
        '        '顯示第一個子視窗       
        '        xlBook.Parent.Windows(1).Visible = True
        '        '引用第一個工作表     
        '        xlSheet = xlBook.Worksheets(1)
        '        '設定工作表為焦點     
        '        xlSheet.Activate()
        '        '================================================================================
        '        Dim count As Double
        '        For i = 0 To DataGridView2.ColumnCount
        '            xlSheet.Cells(1, i + 1) = DataGridView2.Columns(i).HeaderText
        '            count += 1
        '        Next
        '        For i = 0 To DataGridView2.RowCount
        '            For j = 0 To DataGridView2.ColumnCount
        '                xlSheet.Cells(i + 2, j + 1) = "'" + DataGridView2(j, i).Value.ToString
        '                count += 1
        '            Next
        '        Next
        '        'MsgBox("Finish", MsgBoxStyle.SystemModal, "Output Excel")
        '        xlSheet.Cells.EntireColumn.AutoFit()

        '        Me.WindowState = FormWindowState.Minimized
        'endsub2:

        If btnMaxLC.Checked Then
            SaveFileDialog1.FileName = "SideswayCheck(Selected Comb)"
        Else
            SaveFileDialog1.FileName = "SideswayCheck(ALL Comb)"
        End If



        SaveFileDialog1.DefaultExt = "Disp"
        SaveFileDialog1.Title = "Sidesway Check"
        SaveFileDialog1.Filter = "Displement (*.Disp) | *.Disp"
        SaveFileDialog1.InitialDirectory = My.Computer.FileSystem.CurrentDirectory
        'SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then GoTo endsub

        FileOpen(30, SaveFileDialog1.FileName, OpenMode.Output)
        'Dim Header1 As String = " For Column line "

        Dim Header2 As String = "CL     Node      H/" & TextBox1.Text & "     LC          TX1     TX2    |DX|        LC         TY1     TY2    |DY|      LC         |DXY|"
        PrintLine(30, "Unit : mm")
        PrintLine(30, Header2)
        For i = 0 To DataGridView2.RowCount - 1
            If DataGridView2(0, i).Value = "" Then Exit For
            PrintLine(30, DataGridView2(0, i).Value.ToString.Replace(" ", ""), TAB(7), DataGridView2(1, i).Value, TAB(18), DataGridView2(2, i).Value, TAB(27), DataGridView2(3, i).Value, TAB(37), DataGridView2(4, i).Value.ToString.PadLeft(7), TAB(45), DataGridView2(5, i).Value.ToString.PadLeft(7), TAB(53), DataGridView2(6, i).Value.ToString.PadLeft(7), TAB(63), DataGridView2(7, i).Value, TAB(67), DataGridView2(8, i).Value, TAB(75), DataGridView2(9, i).Value.ToString.PadLeft(7), TAB(83), DataGridView2(10, i).Value.ToString.PadLeft(7), TAB(91), DataGridView2(11, i).Value.ToString.PadLeft(7), TAB(100), DataGridView2(12, i).Value, TAB(103), DataGridView2(13, i).Value, TAB(112), DataGridView2(14, i).Value.ToString.PadLeft(7), TAB(120), DataGridView2(15, i).Value.ToString.PadLeft(7))

        Next


        FileClose(30)
        MsgBox("Finish" & vbCrLf & SaveFileDialog1.FileName, , "Sidesway Check")

endsub:

    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            Label1.Text = "Grid Line Data : "
        Else
            Label1.Text = "Result : "
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        'DataGridView2.Columns(2).HeaderText = "H/" & TextBox1.Text & " (mm)"
    End Sub

    Private Sub btnToXML_Click(sender As Object, e As EventArgs) Handles btnToXML.Click

        SaveFileDialog1.DefaultExt = "xml"
        SaveFileDialog1.Title = "Sidesway Check"
        SaveFileDialog1.Filter = "Sidesway Data (*.xml) | *.xml"
        SaveFileDialog1.InitialDirectory = My.Computer.FileSystem.CurrentDirectory
        'SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then GoTo endsub



        Dim settings As XmlWriterSettings = New XmlWriterSettings()
        settings.Indent = True
        Dim NodeString As String
        Dim StartNodeName As String
        Dim EndNodeName As String
        Dim SENodeName() As String

        Dim charseparators() As Char = "-"

        Using writer As XmlWriter = XmlWriter.Create(SaveFileDialog1.FileName, settings)
            ' Begin writing.
            writer.WriteStartDocument()
            writer.WriteStartElement("Sidesway") ' Root.


            For i = 0 To DataGridView2.RowCount - 1
                If DataGridView2(0, i).Value = "" Then Exit For

                SENodeName = DataGridView2(1, i).Value.ToString.Split(charseparators, StringSplitOptions.RemoveEmptyEntries)
                StartNodeName = SENodeName(0).Trim
                EndNodeName = SENodeName(1).Trim

                writer.WriteStartElement("Result")
                writer.WriteElementString("ColumnLine", DataGridView2(0, i).Value.ToString)
                writer.WriteElementString("StartNode", StartNodeName)
                writer.WriteElementString("EndNode", EndNodeName)
                writer.WriteElementString("DispLimitation", DataGridView2(2, i).Value.ToString)
                writer.WriteElementString("XLoadComb", DataGridView2(3, i).Value.ToString)
                writer.WriteElementString("XStartDisp", DataGridView2(4, i).Value.ToString)
                writer.WriteElementString("XEndDisp", DataGridView2(5, i).Value.ToString)
                writer.WriteElementString("XRltvDisp", DataGridView2(6, i).Value.ToString)
                writer.WriteElementString("XIsOK", DataGridView2(7, i).Value.ToString)
                writer.WriteElementString("YLoadComb", DataGridView2(8, i).Value.ToString)
                writer.WriteElementString("YStartDisp", DataGridView2(9, i).Value.ToString)
                writer.WriteElementString("YEndDisp", DataGridView2(10, i).Value.ToString)
                writer.WriteElementString("YRltvDisp", DataGridView2(11, i).Value.ToString)
                writer.WriteElementString("YIsOK", DataGridView2(12, i).Value.ToString)
                writer.WriteElementString("MaxLoadComb", DataGridView2(13, i).Value.ToString)
                writer.WriteElementString("MaxRltvDisp", DataGridView2(14, i).Value.ToString)
                writer.WriteElementString("MaxIsOK", DataGridView2(15, i).Value.ToString)
                writer.WriteEndElement()
            Next

            ' Loop over employees in array.
            'Dim BB As AA
            'For Each BB In DataSet1
            '    writer.WriteStartElement("XDXD")
            '    writer.WriteElementString("ID", BB.Name)
            '    writer.WriteElementString("FirstName", BB.Number.ToString)
            '    writer.WriteElementString("LastName", BB.str)
            '    'writer.WriteElementString("Salary", Employee._salary.ToString)
            '    writer.WriteEndElement()
            'Next

            ' End document.
            writer.WriteEndElement()
            writer.WriteEndDocument()
        End Using

        MsgBox(SaveFileDialog1.FileName, , "Output XML")
endsub:

    End Sub
End Class
