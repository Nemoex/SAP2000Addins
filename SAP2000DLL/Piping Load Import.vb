Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports SAP2000v20
'Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Media.Media3D
Imports System.IO
Imports System.Drawing.Drawing2D
Imports System.Drawing



Public Class Piping_Load_Import
    Public SapModel As SAP2000v20.cSapModel
    Public ISapPlugin As SAP2000v20.cPluginCallback

    Public pipedatacol() As LineData
    Public LineCollection As New Dictionary(Of Integer, LineData)

    '選csv檔 讀取資料
    Private Sub btn_SelectFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_SelectFile.Click

        OpenFileDialog1.Filter = "Piping load data|*.csv"
        OpenFileDialog1.Title = "Select Piping Load Data"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            Try
                FileOpen(10, OpenFileDialog1.FileName, OpenMode.Input)
            Catch ex As Exception
                MsgBox("Error Code 01")
                GoTo endsub
            End Try

            OpenFileName.Text = OpenFileDialog1.FileName

            Dim linetext As String
            Dim dummy As String
            Dim Title As String
            Dim LineDatastock() As String
            Dim linecounter As Integer = 1
            Dim row As String()

            DataGridView1.Rows.Clear()

            Do Until EOF(10)

                linetext = LineInput(10)
                LineDatastock = linetext.Split(",")


                If LineDatastock(0) = "" AndAlso LineDatastock.Count < 2 Then
                    Exit Do
                ElseIf linecounter = 1 Then
                    Title = LineDatastock(3)
                ElseIf linecounter > 2 Then

                    row = New String() {LineDatastock(0), LineDatastock(1), LineDatastock(2), LineDatastock(3), LineDatastock(4), LineDatastock(5), LineDatastock(6), LineDatastock(7), LineDatastock(8), LineDatastock(9), LineDatastock(10), LineDatastock(11), LineDatastock(12), LineDatastock(13), LineDatastock(14), LineDatastock(15), LineDatastock(16)}
                    DataGridView1.Rows.Add(row)
                End If

                linecounter += 1
            Loop
        Else
            MsgBox("Can't Load File")
            GoTo endsub
        End If

endsub:
        FileClose(10)
    End Sub

    '(modify position) 點座標轉置
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        LineCollection.Clear()

        Dim myMatrix As New Matrix()
        myMatrix.Rotate(CDbl(TextBox1.Text))
        myMatrix.Translate(-1 * CDbl(TextBox2.Text), -1 * CDbl(TextBox3.Text), MatrixOrder.Prepend)

        Dim myArray(1) As System.Drawing.PointF
        Dim StartPt, EndPt As New System.Drawing.PointF

        DataGridView2.Rows.Clear()

        Dim row As String()



        For i = 0 To DataGridView1.RowCount - 1

            If DataGridView1.Item(1, i).Value = "" Then GoTo 999

            StartPt = New System.Drawing.PointF(DataGridView1.Item(6, i).Value, DataGridView1(7, i).Value)
            EndPt = New System.Drawing.PointF(DataGridView1.Item(9, i).Value, DataGridView1(10, i).Value)
            myArray(0) = StartPt
            myArray(1) = EndPt
            myMatrix.TransformPoints(myArray)

            'DataGridView2.Item(5, i).Value = StartPt.X
            'DataGridView2.Item(6, i).Value = StartPt.Y
            'DataGridView2.Item(7, i).Value = DataGridView1.Item(7, i).Value - TextBox4.Text

            'DataGridView2.Item(8, i).Value = EndPt.X
            'DataGridView2.Item(9, i).Value = EndPt.Y
            'DataGridView2.Item(10, i).Value = DataGridView1.Item(10, i).Value - TextBox4.Text

            row = New String() {DataGridView1.Item(0, i).Value, DataGridView1.Item(1, i).Value, DataGridView1.Item(2, i).Value, DataGridView1.Item(3, i).Value _
                                , DataGridView1.Item(4, i).Value, DataGridView1.Item(5, i).Value, myArray(0).X.ToString, myArray(0).Y.ToString, DataGridView1.Item(8, i).Value - TextBox4.Text _
                                , myArray(1).X.ToString, myArray(1).Y.ToString, DataGridView1.Item(11, i).Value - TextBox4.Text, DataGridView1.Item(12, i).Value _
                                , DataGridView1.Item(13, i).Value, DataGridView1.Item(14, i).Value, DataGridView1.Item(15, i).Value, DataGridView1.Item(16, i).Value}

            DataGridView2.Rows.Add(row)

            '========產生line Data collection

            Dim tmp As LineData
            tmp.item = row(0)
            tmp.LineNo = row(1)
            tmp.Size = row(2)
            tmp.PipeClass = row(3)
            tmp.Insu_Thk = row(4)
            tmp.OD = row(5)
            tmp.Sx = row(6)
            tmp.Sy = row(7)
            tmp.Sz = row(8)
            tmp.Ex = row(9)
            tmp.Ey = row(10)
            tmp.Ez = row(11)
            tmp.PipeShoe = row(12)
            tmp.PE = row(13)
            tmp.PO = row(14)
            tmp.PT = row(15)
            tmp.Remark = row(16)
            'tmp.needload = False

            LineCollection.Add(tmp.item, tmp)

            '===============================

        Next




999:
        TabControl1.SelectedIndex = 1
        If DataGridView2.Item(0, 0).Value = "" Then
            GroupBox6.Enabled = False
            GroupBox6.Visible = False
        Else
            GroupBox6.Enabled = True
            GroupBox6.Visible = True
        End If



    End Sub


    Private Sub Piping_Load_Import_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Get all Load Case
        Dim ret As Long
        Dim NumberNames As Long
        Dim MyName() As String

        ret = SapModel.LoadPatterns.GetNameList(NumberNames, MyName)

        For i = 0 To NumberNames - 1
            LoadPatterncomb.Items.Add(MyName(i))
        Next

        'Set Units
        LoadPatterncomb.SelectedIndex = 0
        Unitcomb.SelectedIndex = 6    'Units : kgf,mm
        ComboBox3.SelectedIndex = 3   'force direction  : Gravity

    End Sub

    Public Structure Vector
        Public x As Double
        Public y As Double
        '點積運算 沒有除法，儘量避免誤差
        Public Function dot(ByVal v1 As Vector, ByVal v2 As Vector) As Double
            Return v1.x * v2.x + v1.y * v2.y
        End Function
        '叉積運算，回傳純量（除去方向）
        Public Function cross(ByVal v1 As Vector, ByVal v2 As Vector) As Double
            Return v1.x * v2.y - v1.y * v2.x
        End Function
    End Structure   ' Vector

    '向量oa與向量ob進行叉積，判斷oa到ob的旋轉方向。
    Public Function cross(ByRef o As Point, ByRef a As Point, ByRef b As Point) As Double
        Return (a.x - o.x) * (b.y - o.y) - (a.y - o.y) * (b.x - o.x)
    End Function


    Public Structure LineData
        Dim item As Integer
        Dim LineNo As String
        Dim Size As Integer
        Dim PipeClass As String
        Dim Insu_Thk As Integer
        Dim OD As Double
        Dim Sx, Sy, Sz, Ex, Ey, Ez As Double
        Dim PipeShoe As String
        Dim PE, PO, PT As Double
        Dim Remark As String
        Dim needload As String

    End Structure


    Private Structure Member
        Dim Name As String
        Dim SP As StartPoint
        Dim EP As EndPoint
        Dim Length As Double
        Dim BearingLength As Double
        Dim BearingLength_GtoG As Double
        Dim BearingLength_Ext As Double
        Dim isGirder As Boolean
        'Dim BeaingStartLength As Double
        'Dim BearingEndLength As Double
    End Structure

    Private Structure StartPoint
        Dim Name As String
        Dim X, Y, Z As Double
    End Structure

    Private Structure EndPoint
        Dim Name As String
        Dim X, Y, Z As Double
    End Structure

    Public Function getLengthUnit(ByRef Units As String) As String

        Select Case Units
            Case "lb,in,F", "kip,in,F"
                getLengthUnit = "in"
            Case "lb, ft, F", "kip, ft, F"
                getLengthUnit = "ft"
            Case "kN, mm, C", "kgf, mm, C", "N, mm, C", "N, mm, C", "Ton, mm, C"
                getLengthUnit = "mm"
            Case "kN, m, C", "kgf, m, C", "N, m, C", "Ton, m, C"
                getLengthUnit = "m"
            Case "kN, cm, C", "kgf, cm, C", "N, cm, C", "Ton, cm, C"
                getLengthUnit = "cm"
            Case Else
                getLengthUnit = "Unit Error"
        End Select

        Return getLengthUnit

    End Function


    Private Sub btnDirectApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDirectApply.Click


        ''Input check
        'For Each ValueCheck As Control In GroupBox4.Controls
        '    If ValueCheck.GetType Is GetType(System.Windows.Forms.TextBox) Then
        '        If IsNumeric(ValueCheck.Text) = False Then
        '            MsgBox(ValueCheck.Text & " is an illegal value!", MsgBoxStyle.Exclamation, "Regular Point Loads")
        '            GoTo 999
        '        End If
        '    End If
        'Next


        Dim ret As Long
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim LoadPattern As String = LoadPatterncomb.SelectedText
        Dim LoadDirection As String = ComboBox3.SelectedText

        Dim Mytype As Integer
        If RadioButton1.Checked Then Mytype = 1
        If RadioButton2.Checked Then Mytype = 2

        Dim Direction As Integer
        If ComboBox3.SelectedIndex = 0 Then Direction = 4
        If ComboBox3.SelectedIndex = 1 Then Direction = 5
        If ComboBox3.SelectedIndex = 2 Then Direction = 6
        If ComboBox3.SelectedIndex = 3 Then Direction = 10

        Dim Replace As Boolean

        Dim RelDist As Boolean
        If RadioButton6.Checked Then RelDist = True
        If RadioButton7.Checked Then RelDist = False

        '只處理在SAP中最後選的桿件
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        If NumberItems < 2 Then
            MsgBox("Selected member less then 2 , program stop", MsgBoxStyle.Exclamation, "Auto pipe load")
            GoTo 999
        End If

        SapModel.SetPresentUnits(Unitcomb.SelectedIndex + 1)  'SelectedIndex 從0開始  暫時強制單位為Kg / mm


        '==============create load pattern load case

        Dim NumberNames As Long
        Dim MyName() As String

        ret = SapModel.LoadPatterns.GetNameList(NumberNames, MyName)

        If MyName.Contains("AutoPE") And MyName.Contains("AutoPO") And MyName.Contains("AutoPT") Then
            GoTo skipcreatePattern
        End If

        ret = SapModel.LoadPatterns.Add("AutoPE", eLoadPatternType.Other)
        ret = SapModel.LoadPatterns.Add("AutoPO", eLoadPatternType.Other)
        ret = SapModel.LoadPatterns.Add("AutoPT", eLoadPatternType.Other)

        Dim MyLoadType() As String
        Dim MyLoadName() As String
        Dim MySF() As Double
        Dim NumberLoads As Long
        Dim LoadType() As String
        Dim LoadName() As String
        Dim SF() As Double

        'add static linear load case
        ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPE")
        'set load data
        ReDim MyLoadType(0)
        ReDim MyLoadName(0)
        ReDim MySF(0)
        MyLoadType(0) = "Load"
        MyLoadName(0) = "AutoPE"
        MySF(0) = 1
        ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPE", 1, MyLoadType, MyLoadName, MySF)

        'add static linear load case
        ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPO")
        'set load data
        ReDim MyLoadType(0)
        ReDim MyLoadName(0)
        ReDim MySF(0)
        MyLoadType(0) = "Load"
        MyLoadName(0) = "AutoPO"
        MySF(0) = 1
        ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPO", 1, MyLoadType, MyLoadName, MySF)

        'add static linear load case
        ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPT")
        'set load data
        ReDim MyLoadType(0)
        ReDim MyLoadName(0)
        ReDim MySF(0)
        MyLoadType(0) = "Load"
        MyLoadName(0) = "AutoPT"
        MySF(0) = 1
        ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPT", 1, MyLoadType, MyLoadName, MySF)

        '============================================

        '=============================
        'ret = SapModel.LoadCases.StaticNonlinear.SetCase("ASD" + MyName(i))
        'ret = SapModel.RespCombo.Add(MyName(i) + "_ASD", 0)
        'For j = 0 To SF.Count - 1
        '    SF(j) = SF(j) * 1.6
        'Next
        'ret = SapModel.LoadCases.StaticNonlinear.SetLoads("ASD" + MyName(i), NumberLoads, LoadType, LoadName, SF)
        'ret = SapModel.LoadCases.StaticNonlinear.SetGeometricNonlinearity("ASD" + MyName(i), 1)
        'ret = SapModel.RespCombo.SetCaseList(MyName(i) + "_ASD", 0, "ASD" + MyName(i), 0.625)
        '=============================

skipcreatePattern:

        Dim Point1, Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        Dim SelObj As New List(Of Member)     'selected object collection
        Dim m1 As New Member

        Dim SelObj2 As New List(Of Member)    '計算完bearing length的集合
        Dim m2 As New Member

        SapModel.GetPresentUnits()




        For i = 0 To NumberItems - 1
            ret = SapModel.FrameObj.GetPoints(ObjectName(i), Point1, Point2)
            ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
            ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)

            With m1
                .Name = ObjectName(i)
                .SP.Name = Point1
                .SP.X = X1
                .SP.Y = Y1
                .SP.Z = Z1
                .EP.Name = Point2
                .EP.X = X2
                .EP.Y = Y2
                .EP.Z = Z2
                .Length = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5
            End With

            SelObj.Add(m1)

        Next


        '判別所有桿件是否平行

        Dim Vector1, Vector2 As Vector3D
        Vector1.X = SelObj(0).SP.X - SelObj(0).EP.X
        Vector1.Y = SelObj(0).SP.Y - SelObj(0).EP.Y
        Vector1.Z = SelObj(0).SP.Z - SelObj(0).EP.Z

        For i = 1 To NumberItems - 1

            Vector2.X = SelObj(i).SP.X - SelObj(i).EP.X
            Vector2.Y = SelObj(i).SP.Y - SelObj(i).EP.Y
            Vector2.Z = SelObj(i).SP.Z - SelObj(i).EP.Z

            If Vector3D.CrossProduct(Vector1, Vector2).Length > 0.00001 Then
                MsgBox("選的桿件 " & SelObj(i).Name & " 和 " & SelObj(0).Name & " 非平行!")
                GoTo 999
            End If
        Next

        Dim PiperackDirection As String
        If Vector1.X = 0 And Vector1.Y <> 0 Then
            PiperackDirection = "X-Dir"
            SelObj.Sort(Function(p1, p2) p1.SP.X < p2.SP.X)     '在這排序

        ElseIf Vector1.X <> 0 And Vector1.Y = 0 Then
            PiperackDirection = "Y-Dir"
            SelObj.Sort(Function(p1, p2) p1.SP.Y < p2.SP.Y)     '在這排序
        End If

        '====計算選取的桿件其Bearinglength / BearingLength_Ext 並儲存數值

        Dim BearL As Double      'Bearing Length
        Dim BearL_Ext As Double  'Bearing Length Extension

        For i = 0 To SelObj.Count - 1
            If PiperackDirection = "X-Dir" Then
                If i = 0 Then
                    BearL_Ext = CDbl(Bearinglength_Ext.Text)
                    BearL = Math.Abs(SelObj(i).SP.X - SelObj(i + 1).SP.X) / 2
                ElseIf (i = SelObj.Count - 1) Then
                    BearL_Ext = CDbl(Bearinglength_Ext.Text)
                    BearL = Math.Abs(SelObj(i).SP.X - SelObj(i - 1).SP.X) / 2
                Else
                    BearL_Ext = 0
                    BearL = Math.Abs(SelObj(i + 1).SP.X - SelObj(i - 1).SP.X) / 2
                End If

            ElseIf PiperackDirection = "Y-Dir" Then
                If i = 0 Then
                    BearL_Ext = CDbl(Bearinglength_Ext.Text)
                    BearL = Math.Abs(SelObj(i).SP.Y - SelObj(i + 1).SP.Y) / 2
                ElseIf (i = SelObj.Count - 1) Then
                    BearL_Ext = CDbl(Bearinglength_Ext.Text)
                    BearL = Math.Abs(SelObj(i).SP.Y - SelObj(i - 1).SP.Y) / 2
                Else
                    BearL_Ext = 0
                    BearL = Math.Abs(SelObj(i + 1).SP.Y - SelObj(i - 1).SP.Y) / 2
                End If

            End If
            m2 = SelObj(i)
            m2.BearingLength = BearL
            m2.BearingLength_Ext = BearL_Ext
            SelObj2.Add(m2)

        Next

        '============找出選取桿件的主層Z MainElev

        Dim MaxZ, minZ As Double
        MaxZ = -9999999
        minZ = 9999999
        Dim MainElev As Double

        For Each Beam As Member In SelObj
            If Beam.SP.Z > MaxZ Then
                MaxZ = Beam.SP.Z
            End If
            If Beam.SP.Z < minZ Then
                minZ = Beam.SP.Z
            End If
        Next

        If MaxZ - minZ > 500 Then
            MsgBox("Member have different elevation , end program")
            GoTo 999
        Else
            MainElev = (MaxZ + minZ) / 2
        End If

        '============判斷pipe line 是否與member 有交會點,先找elevation 再找XY Range
        Dim pipeElev As Double
        Dim a, b, c, d, _e, f As Double
        Dim N, O As Double


        'Replace 指令只會清除1次原設定,在一開始先做完

        If RadioButton4.Checked Then
            For i = 0 To SelObj2.Count - 1
                ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(i).Name, "AutoPE", eItemType.Objects)
                ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(i).Name, "AutoPO", eItemType.Objects)
                ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(i).Name, "AutoPT", eItemType.Objects)
            Next
        End If
        '============================================



        For i = 0 To DataGridView2.Rows.Count - 1
            pipeElev = DataGridView2.Item(7, i).Value
            If pipeElev > MainElev - 500 And pipeElev < MainElev + 500 Then    '找range +- 主層500 mm 的pipe data

                For j = 0 To SelObj2.Count - 1  '將在主層Range 內的pipe data 與selected member 做有無交點的判斷

                    Dim A1, A2, B1, B2 As Point
                    A1.X = SelObj2(j).SP.X
                    A1.Y = SelObj2(j).SP.Y
                    A2.X = SelObj2(j).EP.X
                    A2.Y = SelObj2(j).EP.Y

                    B1.X = DataGridView2.Item(5, i).Value
                    B1.Y = DataGridView2.Item(6, i).Value
                    B2.X = DataGridView2.Item(8, i).Value
                    B2.Y = DataGridView2.Item(9, i).Value

                    If intersect(A1, A2, B1, B2) = True Then  '有交點則繼續處理 計算pipe與member在何處相交

                        '求解2線交座標點

                        a = A2.Y - A1.Y
                        b = A1.X - A2.X
                        c = a * A1.X + b * A1.Y
                        d = B2.Y - B1.Y
                        _e = B1.X - B2.X
                        f = d * B1.X + _e * B1.Y

                        '交點座標 (N,O)

                        N = (c * _e - b * f) / (a * _e - b * d)
                        O = (c * d - a * f) / (b * d - a * _e)

                        '計算交點與桿件起點的距離

                        Dim distance As Double = 0
                        distance = ((N - SelObj2(j).SP.X) ^ 2 + (O - SelObj2(j).SP.Y) ^ 2) ^ 0.5
                        Dim totalloadlength As Double = (SelObj2(j).BearingLength + SelObj2(j).BearingLength_Ext) / 1000   '管重為 kg / m 進行單位轉換

                        If RadioButton3.Checked Then

                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPE", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(12, i).Value, "Global", RelDist, Replace, eItemType.Objects)
                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPO", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(13, i).Value, "Global", RelDist, Replace, eItemType.Objects)
                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPT", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(14, i).Value, "Global", RelDist, Replace, eItemType.Objects)

                            DataGridView2.Item(16, i).Value = "Add"
                            DataGridView2.Rows(i).Cells(16).Style.BackColor = Color.White

                        End If

                        If RadioButton4.Checked Then

                            'ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPE", eItemType.Object)
                            'ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPO", eItemType.Object)
                            'ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPT", eItemType.Object)

                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPE", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(12, i).Value, "Global", RelDist, Replace, eItemType.Objects)
                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPO", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(13, i).Value, "Global", RelDist, Replace, eItemType.Objects)
                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPT", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(14, i).Value, "Global", RelDist, Replace, eItemType.Objects)

                            DataGridView2.Item(16, i).Value = "Replace"
                            DataGridView2.Rows(i).Cells(16).Style.BackColor = Color.Yellow

                        End If

                        If RadioButton5.Checked = True Then

                            ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPE", eItemType.Objects)
                            ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPO", eItemType.Objects)
                            ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPT", eItemType.Objects)

                            DataGridView2.Item(16, i).Value = "Delete"
                            DataGridView2.Rows(i).Cells(16).Style.BackColor = Color.Red


                        End If

                        'MsgBox("with intersection")
                    Else
                        'MsgBox("no intersection")

                    End If

                Next

            End If

        Next

        MsgBox("Finish", , "Point Loads")
999:

        'Me.Close()
    End Sub

    '檢查2線交會func
    Private Function intersect(ByVal a1 As Point, ByVal a2 As Point, ByVal b1 As Point, ByVal b2 As Point) As Boolean

        Dim c1 As Double = cross(a1, a2, b1)
        Dim c2 As Double = cross(a1, a2, b2)
        Dim c3 As Double = cross(b1, b2, a1)
        Dim c4 As Double = cross(b1, b2, a2)

        ' 端點不共線
        If (c1 * c2 < 0 And c3 * c4 < 0) Then
            Return True
        Else

        End If

    End Function


    '檢查2線交會func
    Private Function CheckIntersection(ByRef Pipedata As LineData, ByRef frame As Member) As Boolean

        If Pipedata.Ex > Pipedata.Sx Then
            If Pipedata.Sx < frame.SP.X And frame.SP.X < Pipedata.Ex Then
                If Pipedata.Sy < 0.0 Then

                    CheckIntersection = True

                End If
            ElseIf Pipedata.Sx > Pipedata.Ex Then
                If Pipedata.Ex < frame.SP.X And frame.SP.X < Pipedata.Sx Then
                    CheckIntersection = True
                End If

            End If

        End If

    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        'Dim ret As Long
        'Dim NumberAreas As Long


        'ret = SapModel.LoadPatterns.Add("AutoPE", eLoadPatternType.LTYPE_OTHER)
        'ret = SapModel.LoadPatterns.Add("AutoPO", eLoadPatternType.LTYPE_OTHER)
        'ret = SapModel.LoadPatterns.Add("AutoPT", eLoadPatternType.LTYPE_OTHER)

        'Dim MyLoadType() As String
        'Dim MyLoadName() As String
        'Dim MySF() As Double
        'Dim NumberLoads As Long
        'Dim LoadType() As String
        'Dim LoadName() As String
        'Dim SF() As Double

        ''add static linear load case
        'ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPE")
        ''set load data
        'ReDim MyLoadType(0)
        'ReDim MyLoadName(0)
        'ReDim MySF(0)
        'MyLoadType(0) = "Load"
        'MyLoadName(0) = "AutoPE"
        'MySF(0) = 1
        'ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPE", 1, MyLoadType, MyLoadName, MySF)

        ''add static linear load case
        'ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPO")
        ''set load data
        'ReDim MyLoadType(0)
        'ReDim MyLoadName(0)
        'ReDim MySF(0)
        'MyLoadType(0) = "Load"
        'MyLoadName(0) = "AutoPO"
        'MySF(0) = 1
        'ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPO", 1, MyLoadType, MyLoadName, MySF)

        ''add static linear load case
        'ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPT")
        ''set load data
        'ReDim MyLoadType(0)
        'ReDim MyLoadName(0)
        'ReDim MySF(0)
        'MyLoadType(0) = "Load"
        'MyLoadName(0) = "AutoPT"
        'MySF(0) = 1
        'ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPT", 1, MyLoadType, MyLoadName, MySF)






    End Sub

    Private Sub btnCheckSpec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckSpec.Click

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboMidBeam.SelectedIndexChanged

    End Sub

    Private Sub btnGirder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGirder.Click

        ComboGirder.Items.Clear()

        Dim ret As Long
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        For i = 0 To NumberItems - 1
            If ObjectType(i) = 2 Then
                ComboGirder.Items.Add(ObjectName(i))
            End If
        Next

    End Sub

    Private Sub btnMidBeam_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMidBeam.Click

        ComboMidBeam.Items.Clear()

        Dim ret As Long
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        For i = 0 To NumberItems - 1
            If ObjectType(i) = 2 Then
                ComboMidBeam.Items.Add(ObjectName(i))
            End If
        Next


    End Sub

    Private Sub btn_HLGirder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_HLGirder.Click
        Dim ret As Long
        ret = SapModel.SelectObj.ClearSelection
        For i = 0 To ComboGirder.Items.Count - 1
            ret = SapModel.FrameObj.SetSelected(ComboGirder.Items(i), True)
        Next
        ret = SapModel.View.RefreshView
    End Sub

    Private Sub btn_HLMidBeam_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_HLMidBeam.Click
        Dim ret As Long
        ret = SapModel.SelectObj.ClearSelection
        For i = 0 To ComboMidBeam.Items.Count - 1
            ret = SapModel.FrameObj.SetSelected(ComboMidBeam.Items(i), True)
        Next
        ret = SapModel.View.RefreshView
    End Sub

    Private Sub btnApplyByType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApplyByType.Click
        Dim ret As Long
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim LoadPattern As String = LoadPatterncomb.SelectedText
        Dim LoadDirection As String = ComboBox3.SelectedText

        Dim Mytype As Integer
        If RadioButton1.Checked Then Mytype = 1
        If RadioButton2.Checked Then Mytype = 2

        Dim Direction As Integer
        If ComboBox3.SelectedIndex = 0 Then Direction = 4
        If ComboBox3.SelectedIndex = 1 Then Direction = 5
        If ComboBox3.SelectedIndex = 2 Then Direction = 6
        If ComboBox3.SelectedIndex = 3 Then Direction = 10

        Dim Replace As Boolean

        Dim RelDist As Boolean
        If RadioButton6.Checked Then RelDist = True
        If RadioButton7.Checked Then RelDist = False

        '===ALL member list 由combolist取得
        'ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        NumberItems = ComboGirder.Items.Count + ComboMidBeam.Items.Count

        ReDim ObjectName(ComboGirder.Items.Count + ComboMidBeam.Items.Count - 1)

        For i = 0 To ComboGirder.Items.Count - 1
            ObjectName(i) = ComboGirder.Items(i).ToString
        Next

        For i = 0 To ComboMidBeam.Items.Count - 1
            ObjectName(ComboGirder.Items.Count + i) = ComboMidBeam.Items(i).ToString
        Next

        '=============================


        If NumberItems < 2 Then
            MsgBox("Selected member less then 2 , program stop", MsgBoxStyle.Exclamation, "Auto pipe load")
            GoTo 999
        End If

        SapModel.SetPresentUnits(Unitcomb.SelectedIndex + 1)  'SelectedIndex 從0開始  暫時強制單位為Kg / mm


        '==============create load pattern load case

        Dim NumberNames As Long
        Dim MyName() As String

        ret = SapModel.LoadPatterns.GetNameList(NumberNames, MyName)

        If MyName.Contains("AutoPE") And MyName.Contains("AutoPO") And MyName.Contains("AutoPT") Then
            GoTo skipcreatePattern
        End If

        ret = SapModel.LoadPatterns.Add("AutoPE", eLoadPatternType.Other)
        ret = SapModel.LoadPatterns.Add("AutoPO", eLoadPatternType.Other)
        ret = SapModel.LoadPatterns.Add("AutoPT", eLoadPatternType.Other)

        Dim MyLoadType() As String
        Dim MyLoadName() As String
        Dim MySF() As Double
        Dim NumberLoads As Long
        Dim LoadType() As String
        Dim LoadName() As String
        Dim SF() As Double

        'add static linear load case
        ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPE")
        'set load data
        ReDim MyLoadType(0)
        ReDim MyLoadName(0)
        ReDim MySF(0)
        MyLoadType(0) = "Load"
        MyLoadName(0) = "AutoPE"
        MySF(0) = 1
        ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPE", 1, MyLoadType, MyLoadName, MySF)

        'add static linear load case
        ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPO")
        'set load data
        ReDim MyLoadType(0)
        ReDim MyLoadName(0)
        ReDim MySF(0)
        MyLoadType(0) = "Load"
        MyLoadName(0) = "AutoPO"
        MySF(0) = 1
        ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPO", 1, MyLoadType, MyLoadName, MySF)

        'add static linear load case
        ret = SapModel.LoadCases.StaticLinear.SetCase("AutoPT")
        'set load data
        ReDim MyLoadType(0)
        ReDim MyLoadName(0)
        ReDim MySF(0)
        MyLoadType(0) = "Load"
        MyLoadName(0) = "AutoPT"
        MySF(0) = 1
        ret = SapModel.LoadCases.StaticLinear.SetLoads("AutoPT", 1, MyLoadType, MyLoadName, MySF)

        '============================================

skipcreatePattern:

        Dim Point1, Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        Dim SelObj As New List(Of Member)     'selected object collection
        Dim m1 As New Member

        Dim SelObj2 As New List(Of Member)    '計算完bearing length的集合
        Dim m2 As New Member

        Dim GirderObj As New List(Of Member)  'Girder 集合

        Dim SelObj3 As New List(Of Member)    '計算完bearing length GtoG 的集合
        Dim m3 As New Member

        Dim mTmp As New Member

        'SapModel.GetPresentUnits()


        For i = 0 To NumberItems - 1
            ret = SapModel.FrameObj.GetPoints(ObjectName(i), Point1, Point2)
            ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
            ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)

            With m1
                .Name = ObjectName(i)
                .SP.Name = Point1
                .SP.X = X1
                .SP.Y = Y1
                .SP.Z = Z1
                .EP.Name = Point2
                .EP.X = X2
                .EP.Y = Y2
                .EP.Z = Z2
                .Length = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5
            End With

            SelObj.Add(m1)

        Next


        '判別所有桿件是否平行

        Dim Vector1, Vector2 As Vector3D
        Vector1.X = SelObj(0).SP.X - SelObj(0).EP.X
        Vector1.Y = SelObj(0).SP.Y - SelObj(0).EP.Y
        Vector1.Z = SelObj(0).SP.Z - SelObj(0).EP.Z

        For i = 1 To NumberItems - 1

            Vector2.X = SelObj(i).SP.X - SelObj(i).EP.X
            Vector2.Y = SelObj(i).SP.Y - SelObj(i).EP.Y
            Vector2.Z = SelObj(i).SP.Z - SelObj(i).EP.Z

            If Vector3D.CrossProduct(Vector1, Vector2).Length > 0.00001 Then
                MsgBox("選的桿件 " & SelObj(i).Name & " 和 " & SelObj(0).Name & " 非平行!")
                GoTo 999
            End If
        Next

        Dim PiperackDirection As String
        If Vector1.X = 0 And Vector1.Y <> 0 Then
            PiperackDirection = "X-Dir"
            SelObj.Sort(Function(p1, p2) p1.SP.X < p2.SP.X)     '在這排序

        ElseIf Vector1.X <> 0 And Vector1.Y = 0 Then
            PiperackDirection = "Y-Dir"
            SelObj.Sort(Function(p1, p2) p1.SP.Y < p2.SP.Y)     '在這排序
        End If


        '如果是girder 則標註
        '=================================================================
        For j = 0 To SelObj.Count - 1
            If ComboGirder.Items.Contains(SelObj(j).Name) Then
                mTmp = SelObj(j)
                mTmp.isGirder = True
                SelObj(j) = mTmp
                GirderObj.Add(mTmp)


            End If
        Next

        '====計算選取的桿件其Bearinglength / BearingLength_Ext 並儲存數值

        Dim BearL As Double      'Bearing Length
        Dim BearL_Ext As Double  'Bearing Length Extension


        For i = 0 To SelObj.Count - 1
            If PiperackDirection = "X-Dir" Then
                If i = 0 Then
                    BearL_Ext = CDbl(Bearinglength_Ext.Text)
                    BearL = Math.Abs(SelObj(i).SP.X - SelObj(i + 1).SP.X) / 2
                ElseIf (i = SelObj.Count - 1) Then
                    BearL_Ext = CDbl(Bearinglength_Ext.Text)
                    BearL = Math.Abs(SelObj(i).SP.X - SelObj(i - 1).SP.X) / 2
                Else
                    BearL_Ext = 0
                    BearL = Math.Abs(SelObj(i + 1).SP.X - SelObj(i - 1).SP.X) / 2
                End If

            ElseIf PiperackDirection = "Y-Dir" Then
                If i = 0 Then
                    BearL_Ext = CDbl(Bearinglength_Ext.Text)
                    BearL = Math.Abs(SelObj(i).SP.Y - SelObj(i + 1).SP.Y) / 2
                ElseIf (i = SelObj.Count - 1) Then
                    BearL_Ext = CDbl(Bearinglength_Ext.Text)
                    BearL = Math.Abs(SelObj(i).SP.Y - SelObj(i - 1).SP.Y) / 2
                Else
                    BearL_Ext = 0
                    BearL = Math.Abs(SelObj(i + 1).SP.Y - SelObj(i - 1).SP.Y) / 2
                End If

            End If

            m2 = SelObj(i)
            m2.BearingLength = BearL
            m2.BearingLength_Ext = BearL_Ext
            SelObj2.Add(m2)


        Next

        '========計算選取的Girder其Bearinglength_GtoG並儲存數值
        Dim BearL_GtG As Double    'Bearing Length Girder to Girder
        Dim GirderCount As Integer = 0

        For i = 0 To SelObj2.Count - 1

            BearL_GtG = 0
            If SelObj2(i).isGirder = True Then
                If PiperackDirection = "X-Dir" Then
                    If i = 0 Then
                        BearL_GtG = Math.Abs(GirderObj(GirderCount).SP.X - GirderObj(GirderCount + 1).SP.X) / 2
                    ElseIf (i = SelObj.Count - 1) Or GirderCount >= ComboGirder.Items.Count - 1 Then
                        BearL_GtG = Math.Abs(GirderObj(GirderCount).SP.X - GirderObj(GirderCount - 1).SP.X) / 2
                    Else
                        BearL_GtG = Math.Abs(GirderObj(GirderCount + 1).SP.X - GirderObj(GirderCount - 1).SP.X) / 2
                    End If

                ElseIf PiperackDirection = "Y-Dir" Then
                    If i = 0 Then
                        BearL_GtG = Math.Abs(GirderObj(GirderCount).SP.Y - GirderObj(GirderCount + 1).SP.Y) / 2
                    ElseIf (i = SelObj.Count - 1) Or GirderCount >= ComboGirder.Items.Count - 1 Then
                        BearL_GtG = Math.Abs(GirderObj(GirderCount).SP.Y - GirderObj(GirderCount - 1).SP.Y) / 2
                    Else
                        BearL_GtG = Math.Abs(GirderObj(GirderCount + 1).SP.Y - GirderObj(GirderCount - 1).SP.Y) / 2
                    End If

                End If
                GirderCount += 1
            End If

            mTmp = SelObj2(i)
            mTmp.BearingLength_GtoG = BearL_GtG
            SelObj3.Add(mTmp)
        Next

        '11/02
        '============找出選取桿件的主層Z MainElev

        Dim MaxZ, minZ As Double
        MaxZ = -9999999
        minZ = 9999999
        Dim MainElev As Double

        For Each Beam As Member In SelObj
            If Beam.SP.Z > MaxZ Then
                MaxZ = Beam.SP.Z
            End If
            If Beam.SP.Z < minZ Then
                minZ = Beam.SP.Z
            End If
        Next

        If MaxZ - minZ > 500 Then
            MsgBox("Member have different elevation , end program")
            GoTo 999
        Else
            MainElev = (MaxZ + minZ) / 2
        End If

        '============判斷pipe line 是否與member 有交會點,先找elevation 再找XY Range
        Dim pipeElev As Double
        Dim a, b, c, d, _e, f As Double
        Dim N, O As Double


        'Replace 指令只會清除1次原設定,在一開始先做完

        If RadioButton4.Checked Then
            For i = 0 To SelObj2.Count - 1
                ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(i).Name, "AutoPE", eItemType.Objects)
                ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(i).Name, "AutoPO", eItemType.Objects)
                ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(i).Name, "AutoPT", eItemType.Objects)
            Next
        End If
        '============================================

        Dim pipeDiaSummary As Double = 0
        Dim pipeNumber As Integer = 0
        Dim AvgpipeDia As Double = 0

        '======='計算平均管徑=======
        For i = 0 To DataGridView2.Rows.Count - 1
            pipeElev = DataGridView2.Item(8, i).Value
            If pipeElev > MainElev - 500 And pipeElev < MainElev + 500 Then
                pipeDiaSummary = pipeDiaSummary + DataGridView2.Item(2, i).Value
                pipeNumber += 1
            End If
        Next
        AvgpipeDia = pipeDiaSummary / pipeNumber
        '===========================
        Dim bigpipecount As Integer

        For i = 0 To DataGridView2.Rows.Count - 1
            pipeElev = DataGridView2.Item(8, i).Value
            If pipeElev > MainElev - 500 And pipeElev < MainElev + 500 Then
                If DataGridView2.Item(2, i).Value > AvgpipeDia + 100 Then
                    bigpipecount += 1

                End If
            End If
        Next



        For i = 0 To DataGridView2.Rows.Count - 1
            pipeElev = DataGridView2.Item(8, i).Value
            If pipeElev > MainElev - 500 And pipeElev < MainElev + 500 Then    '找range +- 主層500 mm 的pipe data

                For j = 0 To SelObj2.Count - 1  '將在主層Range 內的pipe data 與selected member 做有無交點的判斷

                    Dim A1, A2, B1, B2 As Point
                    A1.X = SelObj2(j).SP.X
                    A1.Y = SelObj2(j).SP.Y
                    A2.X = SelObj2(j).EP.X
                    A2.Y = SelObj2(j).EP.Y

                    B1.X = DataGridView2.Item(6, i).Value
                    B1.Y = DataGridView2.Item(7, i).Value
                    B2.X = DataGridView2.Item(9, i).Value
                    B2.Y = DataGridView2.Item(10, i).Value

                    If intersect(A1, A2, B1, B2) = True Then  '有交點則繼續處理 計算pipe與member在何處相交

                        '求解2線交座標點

                        a = A2.Y - A1.Y
                        b = A1.X - A2.X
                        c = a * A1.X + b * A1.Y
                        d = B2.Y - B1.Y
                        _e = B1.X - B2.X
                        f = d * B1.X + _e * B1.Y

                        '交點座標 (N,O)

                        N = (c * _e - b * f) / (a * _e - b * d)
                        O = (c * d - a * f) / (b * d - a * _e)

                        '計算交點與桿件起點的距離

                        Dim distance As Double = 0
                        distance = ((N - SelObj2(j).SP.X) ^ 2 + (O - SelObj2(j).SP.Y) ^ 2) ^ 0.5
                        Dim totalloadlength As Double = (SelObj2(j).BearingLength + SelObj2(j).BearingLength_Ext) / 1000   '管重為 kg / m 進行單位轉換

                        If RadioButton3.Checked Then

                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPE", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(13, i).Value, "Global", RelDist, Replace, eItemType.Objects)
                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPO", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(14, i).Value, "Global", RelDist, Replace, eItemType.Objects)
                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPT", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(15, i).Value, "Global", RelDist, Replace, eItemType.Objects)

                            DataGridView2.Item(17, i).Value = "Add"
                            DataGridView2.Rows(i).Cells(17).Style.BackColor = Color.White

                        End If

                        If RadioButton4.Checked Then

                            'ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPE", eItemType.Object)
                            'ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPO", eItemType.Object)
                            'ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPT", eItemType.Object)

                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPE", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(13, i).Value, "Global", RelDist, Replace, eItemType.Objects)
                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPO", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(14, i).Value, "Global", RelDist, Replace, eItemType.Objects)
                            ret = SapModel.FrameObj.SetLoadPoint(SelObj(j).Name, "AutoPT", Mytype, Direction, distance, totalloadlength * DataGridView2.Item(15, i).Value, "Global", RelDist, Replace, eItemType.Objects)

                            DataGridView2.Item(17, i).Value = "Replace"
                            DataGridView2.Rows(i).Cells(17).Style.BackColor = Color.Yellow

                        End If

                        If RadioButton5.Checked = True Then

                            ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPE", eItemType.Objects)
                            ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPO", eItemType.Objects)
                            ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(j).Name, "AutoPT", eItemType.Objects)

                            DataGridView2.Item(16, i).Value = "Delete"
                            DataGridView2.Rows(i).Cells(16).Style.BackColor = Color.Red


                        End If

                        'MsgBox("with intersection")
                    Else
                        'MsgBox("no intersection")

                    End If

                Next

            End If

        Next

        '============計算LOADING 距離桿件起點位置





        '==========找出DataGrid elevation 滿足Mainelev 的資料列
        'Dim usedLength As String
        'usedLength = getLengthUnit(Unitcomb.SelectedText)
        'If usedLength = "m" Then
        'ElseIf usedLength = "mm" Then
        'End If


        Dim key As Integer
        Dim val As LineData
        Dim loadcounter As Integer

        For Each InputData In LineCollection


            key = InputData.Key
            val = InputData.Value

            If val.Sz < MainElev + 500 And val.Sz > MainElev - 500 Then
                If val.Sx > 1 Then

                End If

                'val.needload = "Y"
                'DataGridView2.Item(16, key - 1).Value = "Y"
                'loadcounter += 1


            End If

        Next



        MsgBox("Finish", , "Point Loads")
999:

    End Sub
End Class

