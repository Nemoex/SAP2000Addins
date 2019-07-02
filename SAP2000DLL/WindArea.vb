Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports SAP2000v20
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Media.Media3D
Imports System.Runtime.InteropServices

Public Class WindArea

    Public SapModel As SAP2000v20.cSapModel
    Public ISapPlugin As SAP2000v20.cPluginCallback



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        DataGridView1.Rows.Clear()

        Dim ret As Long
        Dim ParallelTo(5) As Boolean
        Dim Num As Long
        Dim NewName() As String
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim memberCounter, NewMember As Integer

        Dim NumberResults As Integer
        Dim FrameName() As String
        Dim Ratio() As Double
        Dim RatioType() As Integer
        Dim Location() As Double
        Dim ComboName() As String
        Dim ErrorSummary() As String
        Dim WarningSummary() As String

        Dim PropName As String
        Dim BeginSection As String
        Dim DesignSection As String
        Dim SAuto As String

        'get all points position
        Dim pointcount As Integer
        pointcount = SapModel.PointObj.Count


        ret = SapModel.SetPresentUnits(eUnits.Ton_m_C)


        'ret = SapModel.SelectObj.ClearSelection
        'ret = SapModel.SelectObj.All

        Dim SectionName As String
        Dim Start_Z, End_Z As Double
        Dim Length, Ang As Double
        Dim PWidth, PArea As Double


        Dim Rows As New List(Of String())


        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        Dim counter As Integer

        For i = 0 To NumberItems - 1
            If ObjectType(i) = 2 Then
                getFrameData(ObjectName(i), SectionName, Start_Z, End_Z, Length, Ang, PWidth, PArea)
                Dim row As String()
                row = New String() {ObjectName(i), SectionName, Length, Start_Z, End_Z, Ang, PWidth, PArea}
                Rows.Add(row)
                'DataGridView1.Rows.Add(row)
                counter += 1
            End If
        Next

        For j = 0 To Rows.Count - 1
            DataGridView1.Rows.Add(Rows(j))
        Next

    End Sub

    Private Sub WindArea_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Function getFrameData(ByVal FrameNum As String, ByRef SectionName As String, ByRef Start_Z As Double, ByRef End_Z As Double, ByRef Length As Double, ByRef Ang As Double, ByRef PWidth As Double, ByRef PArea As Double) As Boolean
        Dim ret As Long
        Dim SAuto As String
        Dim Area As Double
        Dim as2 As Double
        Dim as3 As Double
        Dim Torsion As Double
        Dim I22 As Double
        Dim I33 As Double
        Dim S22 As Double
        Dim S33 As Double
        Dim Z22 As Double
        Dim Z33 As Double
        Dim R22 As Double
        Dim R33 As Double

        'ret = SapModel.PropFrame.GetSectProps(SectionName, Area, as2, as3, Torsion, I22, I33, S22, S33, Z22, Z33, R22, R33)

        ret = SapModel.FrameObj.GetSection(FrameNum, SectionName, SAuto)

        Dim Point1, Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        ret = SapModel.FrameObj.GetPoints(FrameNum, Point1, Point2)

        'Dim Pt1, Pt2 As New Points_XYZ

        'Pt1 = Points_Position.Item(Point1)
        'Pt2 = Points_Position.Item(Point2)

        ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
        ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)

        Start_Z = Z1
        End_Z = Z2

        Dim Start_X = X1
        Dim End_X = X2

        Dim Start_Y = Y1
        Dim End_Y = Y2


        Length = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5

        Dim Advanced As Boolean
        ret = SapModel.FrameObj.GetLocalAxes(FrameNum, Ang, Advanced)



        ''==================
        'Dim MatProp As String

        'MatProp = Section_material(SectionName)

        '==================
        Dim NameInFile As String
        Dim FileName As String
        Dim MatProp As String
        Dim PropType As eFramePropType

        Dim t3 As Double
        Dim t2 As Double
        Dim tf As Double
        Dim tw As Double
        Dim t2b As Double
        Dim tfb As Double
        Dim Color As Long
        Dim Notes As String
        Dim GUID As String
        Dim dis As Double
        Dim Thickness As Double
        Dim Radius As Double
        Dim LipDepth As Double
        Dim LipAngle As Double

        Dim NumberItems As Long
        Dim ShapeName() As String
        Dim MyType() As Integer
        Dim DesignType As Long


        ret = SapModel.PropFrame.GetTypeOAPI(SectionName, PropType)
        Select Case PropType
            Case eFramePropType.I
                ret = SapModel.PropFrame.GetISection(SectionName, FileName, MatProp, t3, t2, tf, tw, t2b, tfb, Color, Notes, GUID)
            Case eFramePropType.Channel
                ret = SapModel.PropFrame.GetChannel(SectionName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
            Case eFramePropType.T
                ret = SapModel.PropFrame.GetTee(SectionName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
            Case eFramePropType.Angle
                ret = SapModel.PropFrame.GetAngle(SectionName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
            Case eFramePropType.DblAngle
                ret = SapModel.PropFrame.GetDblAngle(SectionName, FileName, MatProp, t3, t2, tf, tw, dis, Color, Notes, GUID)
            Case eFramePropType.Box

            Case eFramePropType.Pipe
                ret = SapModel.PropFrame.GetPipe(SectionName, FileName, MatProp, t3, tw, Color, Notes, GUID)
            Case eFramePropType.Rectangular
                ret = SapModel.PropFrame.GetRectangle(SectionName, FileName, MatProp, t3, t2, Color, Notes, GUID)
            Case eFramePropType.Circle
                ret = SapModel.PropFrame.GetCircle(SectionName, FileName, MatProp, t3, Color, Notes, GUID)
            Case eFramePropType.General
                ret = SapModel.PropFrame.GetGeneral(SectionName, FileName, MatProp, t3, t2, Area, as2, as3, Torsion, I22, I33, S22, S33, Z22, Z33, R22, R33, Color, Notes, GUID)
            Case eFramePropType.DbChannel
                ret = SapModel.PropFrame.GetDblChannel(SectionName, FileName, MatProp, t3, t2, tf, tw, dis, Color, Notes, GUID)
            Case eFramePropType.Auto

            Case eFramePropType.SD
                ret = SapModel.PropFrame.GetSDSection(SectionName, MatProp, NumberItems, ShapeName, MyType, DesignType, Color, Notes, GUID)
            Case eFramePropType.Variable

            Case eFramePropType.Joist

            Case eFramePropType.Bridge

            Case eFramePropType.Cold_C
                ret = SapModel.PropFrame.GetColdC(SectionName, FileName, MatProp, t3, t2, Thickness, Radius, LipDepth, Color, Notes, GUID)
            Case eFramePropType.Cold_2C

            Case eFramePropType.Cold_Z
                ret = SapModel.PropFrame.GetColdZ(SectionName, FileName, MatProp, t3, t2, Thickness, Radius, LipDepth, LipAngle, Color, Notes, GUID)
            Case eFramePropType.Cold_L

            Case eFramePropType.Cold_2L

            Case eFramePropType.Cold_Hat
                ret = SapModel.PropFrame.GetColdHat(SectionName, FileName, MatProp, t3, t2, Thickness, Radius, LipDepth, Color, Notes, GUID)
            Case eFramePropType.BuiltupICoverplate

            Case eFramePropType.PCCGirderI
                'RC
            Case eFramePropType.PCCGirderU
                'RC
        End Select


        '===以Z值差異判斷桿件方向 Z > 1.5 m 視為垂直方向桿件
        Dim VerticalMember As Boolean = False
        If Math.Abs(Z1 - Z2) > 1.5 Then
            VerticalMember = True
        End If
        '========================

        If RadioButton2.Checked Then    '風向Y向
            If VerticalMember = True Then   '垂直桿件
                If Math.Cos(Ang * Math.PI / 180) < 0.00001 Then
                    PWidth = t2
                    PArea = PWidth * Length
                Else
                    PWidth = t3 * Math.Cos(Ang * Math.PI / 180)
                    PArea = PWidth * Length
                End If
            End If

            If VerticalMember = False Then   '水平桿件
                If Math.Cos(Ang * Math.PI / 180) < 0.00001 Then
                    PWidth = t2
                    PArea = PWidth * Length
                Else
                    PWidth = t3 * Math.Cos(Ang * Math.PI / 180)
                    PArea = PWidth * Length
                End If
                If Math.Abs(Start_X - End_X) < 0.001 Then     '桿件沿Y軸，投影面積為0
                    PWidth = PWidth * 0
                    PArea = PArea * 0
                End If

            End If

        ElseIf RadioButton1.Checked Then       '風向X向
            If VerticalMember = True Then    '垂直桿件(柱)
                If Math.Cos(Ang * Math.PI / 180) < 0.00001 Then
                    PWidth = t3
                    PArea = PWidth * Length
                Else
                    PWidth = t2 * Math.Cos(Ang * Math.PI / 180)
                    PArea = PWidth * Length
                End If
            End If

            If VerticalMember = False Then   '水平桿件(梁)
                If Math.Cos(Ang * Math.PI / 180) < 0.00001 Then
                    PWidth = t2
                    PArea = PWidth * Length
                Else
                    PWidth = t3 * Math.Cos(Ang * Math.PI / 180)
                    PArea = PWidth * Length
                End If
                If Math.Abs(Start_Y - End_Y) < 0.001 Then     '桿件沿X軸，投影面積為0
                    PWidth = PWidth * 0
                    PArea = PArea * 0
                End If

            End If
        End If

    End Function


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If DataGridView1.RowCount <= 1 Then GoTo endsub2

        On Error Resume Next
        xlApp = CreateObject("Excel.Application")
        xlBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value)

        '設置EXCEL對象可見        
        xlApp.Visible = True
        '設定活頁簿為焦點        
        xlBook.Activate()
        '顯示第一個子視窗       
        xlBook.Parent.Windows(1).Visible = True
        '引用第一個工作表     
        xlSheet = xlBook.Worksheets(1)
        '設定工作表為焦點     
        xlSheet.Activate()
        '================================================================================
        Dim count As Double

        'xlSheet.Cells(1, 1) = "Unit : Meter"

        For i = 0 To DataGridView1.ColumnCount
            xlSheet.Cells(1, i + 1) = DataGridView1.Columns(i).HeaderText
            count += 1
        Next

        For i = 0 To DataGridView1.RowCount
            For j = 0 To DataGridView1.ColumnCount
                xlSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString
                count += 1
            Next
        Next

        xlSheet.Columns.AutoFit()

        MsgBox("Finish", MsgBoxStyle.SystemModal, "Output Excel")

endsub2:
    End Sub
End Class