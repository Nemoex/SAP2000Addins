Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Media.Media3D
Imports SAP2000v20
Imports System.IO

Public Class Piperack_Distributed_Loads

    Public SapModel As SAP2000v20.cSapModel
    Public ISapPlugin As SAP2000v20.cPluginCallback

    Private Sub Piperack_Point_Loads_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Get all Load Case
        Dim ret As Long
        Dim NumberNames As Integer
        Dim MyName() As String

        ret = SapModel.LoadPatterns.GetNameList(NumberNames, MyName)

        For i = 0 To NumberNames - 1
            LoadPatterncomb.Items.Add(MyName(i))
        Next

        'Set Units
        LoadPatterncomb.SelectedIndex = 0
        Unitcomb.SelectedIndex = 11
        ComboBox3.SelectedIndex = 3

        readSelectionRecord()


    End Sub

    Private Structure Member
        Dim Name As String
        Dim SP As StartPoint
        Dim EP As EndPoint
        Dim Length As Double
    End Structure
    Private Structure StartPoint
        Dim Name As String
        Dim X, Y, Z As Double
    End Structure

    Private Structure EndPoint
        Dim Name As String
        Dim X, Y, Z As Double
    End Structure


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Input check
        For Each ValueCheck As Control In GroupBox4.Controls
            If ValueCheck.GetType Is GetType(System.Windows.Forms.TextBox) Then
                If IsNumeric(ValueCheck.Text) = False Then
                    MsgBox(ValueCheck.Text & " is an illegal value!", MsgBoxStyle.Exclamation, "Regular Point Loads")
                    GoTo 999
                End If
            End If
        Next


        Dim ret As Long
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim LoadPattern As String = LoadPatterncomb.SelectedText
        Dim LoadDirection As String = ComboBox3.SelectedText

        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
        SapModel.SetPresentUnits(Unitcomb.SelectedIndex + 1)  'SelectedIndex 從0開始

        If NumberItems <= 1 Then
            MsgBox("Must select at least 2 members")
            GoTo 999
        End If



        Dim Point1, Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        Dim SelObj As New List(Of Member)

        Dim m1 As New Member

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
        'SelObj.Sort(Function(p1, p2) p1.SP.X < p2.SP.X)

        '判別所有桿件是否平行

        Dim Vector1, Vector2 As Vector3D
        Vector1.X = SelObj(0).SP.X - SelObj(0).EP.X
        Vector1.Y = SelObj(0).SP.Y - SelObj(0).EP.Y
        Vector1.Z = SelObj(0).SP.Z - SelObj(0).EP.Z

        For i = 1 To NumberItems - 1

            Vector2.X = SelObj(i).SP.X - SelObj(i).EP.X
            Vector2.Y = SelObj(i).SP.Y - SelObj(i).EP.Y
            Vector2.Z = SelObj(i).SP.Z - SelObj(i).EP.Z

            If Vector3D.CrossProduct(Vector1, Vector2).Length <> 0 Then
                MsgBox("選的桿件 " & SelObj(i).Name & " 和 " & SelObj(0).Name & " 非平行!")
                GoTo 999
            End If
        Next

        Dim PiperackDirection As String
        If Vector1.X = 0 And Vector1.Y <> 0 Then
            PiperackDirection = "X-Dir"
            SelObj.Sort(Function(p1, p2) p1.SP.X < p2.SP.X)
        ElseIf Vector1.X <> 0 And Vector1.Y = 0 Then
            PiperackDirection = "Y-Dir"
            SelObj.Sort(Function(p1, p2) p1.SP.Y < p2.SP.Y)
        End If

        '計算各桿件受到的Loading長度
        Dim loadlength As Double
        'Dim Value() As Double

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


        Dim Distance(3) As Double
        Distance(0) = CDbl(Distance1.Text)
        Distance(1) = CDbl(Distance2.Text)
        Distance(2) = CDbl(Distance3.Text)
        Distance(3) = CDbl(Distance4.Text)

        Dim Load(3) As Double
        Load(0) = CDbl(Load1.Text)
        Load(1) = CDbl(Load2.Text)
        Load(2) = CDbl(Load3.Text)
        Load(3) = CDbl(Load4.Text)


        For i = 0 To NumberItems - 1
            If RadioButton3.Checked Then Replace = False
            If RadioButton4.Checked Then Replace = True

            If PiperackDirection = "X-Dir" Then
                If i = 0 Then
                    loadlength = CDbl(beginLength.Text) + Math.Abs(SelObj(i).SP.X - SelObj(i + 1).SP.X) / 2
                ElseIf (i = NumberItems - 1) Then
                    loadlength = Math.Abs(SelObj(i).SP.X - SelObj(i - 1).SP.X) / 2 + CDbl(Endlength.Text)
                Else
                    loadlength = Math.Abs(SelObj(i + 1).SP.X - SelObj(i - 1).SP.X) / 2
                End If

                If Replace = True Then
                    ret = SapModel.FrameObj.DeleteLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, eItemType.Objects)
                    Replace = False
                End If

                For j = 0 To 2
                    If Distance(j + 1) > Distance(j) Then
                        ret = SapModel.FrameObj.SetLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, Mytype, Direction, Distance(j), Distance(j + 1), loadlength * Load(j), loadlength * Load(j + 1), "Global", RelDist, Replace, eItemType.Objects)
                    End If
                Next

                'ret = SapModel.FrameObj.SetLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, Mytype, Direction, Distance1.Text, Distance2.Text, loadlength * Load1.Text, loadlength * Load2.Text, "Global", RelDist, Replace, eItemType.Object)
                'ret = SapModel.FrameObj.SetLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, Mytype, Direction, Distance2.Text, Distance3.Text, loadlength * Load2.Text, loadlength * Load3.Text, "Global", RelDist, Replace, eItemType.Object)
                'ret = SapModel.FrameObj.SetLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, Mytype, Direction, Distance3.Text, Distance4.Text, loadlength * Load3.Text, loadlength * Load4.Text, "Global", RelDist, Replace, eItemType.Object)

                If RadioButton5.Checked = True Then
                    ret = SapModel.FrameObj.DeleteLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, eItemType.Objects)
                End If

            ElseIf PiperackDirection = "Y-Dir" Then
                If i = 0 Then
                    loadlength = CDbl(beginLength.Text) + Math.Abs(SelObj(i).SP.Y - SelObj(i + 1).SP.Y) / 2
                ElseIf (i = NumberItems - 1) Then
                    loadlength = Math.Abs(SelObj(i).SP.Y - SelObj(i - 1).SP.Y) / 2 + CDbl(Endlength.Text)
                Else
                    loadlength = Math.Abs(SelObj(i + 1).SP.Y - SelObj(i - 1).SP.Y) / 2
                End If

                If Replace = True Then
                    ret = SapModel.FrameObj.DeleteLoadPoint(SelObj(i).Name, LoadPatterncomb.Text, eItemType.Objects)
                    Replace = False
                End If

                For j = 0 To 2
                    If Distance(j + 1) > Distance(j) Then
                        ret = SapModel.FrameObj.SetLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, Mytype, Direction, Distance(j), Distance(j + 1), loadlength * Load1.Text, loadlength * Load2.Text, "Global", RelDist, Replace, eItemType.Objects)
                    End If
                Next

                'ret = SapModel.FrameObj.SetLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, Mytype, Direction, Distance1.Text, Distance2.Text, loadlength * Load1.Text, loadlength * Load2.Text, "Global", RelDist, Replace, eItemType.Object)
                'ret = SapModel.FrameObj.SetLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, Mytype, Direction, Distance2.Text, Distance3.Text, loadlength * Load2.Text, loadlength * Load3.Text, "Global", RelDist, Replace, eItemType.Object)
                'ret = SapModel.FrameObj.SetLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, Mytype, Direction, Distance3.Text, Distance4.Text, loadlength * Load3.Text, loadlength * Load4.Text, "Global", RelDist, Replace, eItemType.Object)

                If RadioButton5.Checked = True Then
                    ret = SapModel.FrameObj.DeleteLoadDistributed(SelObj(i).Name, LoadPatterncomb.Text, eItemType.Objects)
                End If

            End If
        Next

        writeSelectionRecord()
        MsgBox("Finish", , "Distributed Loads")

999:

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton7.Checked = True Then
            Distance1.Text = "0"
            Distance2.Text = "0"
            Distance3.Text = "0"
            Distance4.Text = "0"
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton6.Checked = True Then
            Distance1.Text = "0"
            Distance2.Text = "0.25"
            Distance3.Text = "0.75"
            Distance4.Text = "1"
        End If
    End Sub

    Private Sub writeSelectionRecord()

        'Dim sw As StreamWriter = New StreamWriter("C:\SAPaddinTemp1.ini", False, System.Text.Encoding.Default)

        Dim sw As StreamWriter = New StreamWriter("D:\SAPaddinTemp1.ini", False, System.Text.Encoding.Default)


        sw.WriteLine(LoadPatterncomb.Text)
        sw.WriteLine(Unitcomb.Text)

        Dim selection1 As Integer
        If RadioButton1.CheckAlign Then
            selection1 = 1
        Else
            selection1 = 2
        End If
        sw.WriteLine(selection1)

        sw.WriteLine(ComboBox3.Text)

        Dim selection2 As Integer
        If RadioButton3.Checked Then
            selection2 = 1
        ElseIf RadioButton4.Checked Then
            selection2 = 2
        Else
            selection2 = 3
        End If

        sw.WriteLine(selection2)

        sw.WriteLine(Distance1.Text)
        sw.WriteLine(Distance2.Text)
        sw.WriteLine(Distance3.Text)
        sw.WriteLine(Distance4.Text)
        sw.WriteLine(Load1.Text)
        sw.WriteLine(Load2.Text)
        sw.WriteLine(Load3.Text)
        sw.WriteLine(Load4.Text)

        sw.WriteLine(beginLength.Text)
        sw.WriteLine(Endlength.Text)

        Dim selection3 As Integer
        If RadioButton6.Checked Then
            selection3 = 1
        Else
            selection3 = 2
        End If



        sw.WriteLine(selection3)

        sw.Close()



    End Sub

    Private Sub readSelectionRecord()
        If My.Computer.FileSystem.FileExists("D:\SAPAddinTemp1.ini") Then
            Dim sr As StreamReader = New StreamReader("D:\SAPAddinTemp1.ini")

            LoadPatterncomb.Text = sr.ReadLine
            Unitcomb.Text = sr.ReadLine

            Dim selection1 As Integer
            selection1 = sr.ReadLine
            If selection1 = 1 Then
                RadioButton1.Checked = True
                RadioButton2.Checked = False
            Else
                RadioButton1.Checked = False
                RadioButton2.Checked = True
            End If

            ComboBox3.Text = sr.ReadLine

            Dim selection2 As Integer
            selection2 = sr.ReadLine
            If selection2 = 1 Then
                RadioButton3.Checked = True
                RadioButton4.Checked = False
                RadioButton5.Checked = False
            ElseIf selection2 = 2 Then
                RadioButton3.Checked = False
                RadioButton4.Checked = True
                RadioButton5.Checked = False
            Else
                RadioButton3.Checked = False
                RadioButton4.Checked = False
                RadioButton5.Checked = True
            End If

            Distance1.Text = sr.ReadLine
            Distance2.Text = sr.ReadLine
            Distance3.Text = sr.ReadLine
            Distance4.Text = sr.ReadLine
            Load1.Text = sr.ReadLine
            Load2.Text = sr.ReadLine
            Load3.Text = sr.ReadLine
            Load4.Text = sr.ReadLine
            beginLength.Text = sr.ReadLine
            Endlength.Text = sr.ReadLine

            Dim selection3 As Integer
            selection3 = sr.ReadLine

            If selection3 = 1 Then
                RadioButton6.Checked = True
                RadioButton7.Checked = False
            Else
                RadioButton6.Checked = False
                RadioButton7.Checked = True
            End If


            sr.Close()
        End If
    End Sub


    Private Sub GroupBox6_Enter(sender As Object, e As EventArgs) Handles GroupBox6.Enter

    End Sub
End Class
