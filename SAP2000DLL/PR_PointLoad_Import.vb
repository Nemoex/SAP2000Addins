Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports SAP2000v20
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Media.Media3D
Imports System.IO

Public Class PR_PointLoad_Import

    Public SapModel As SAP2000v20.cSapModel
    Public ISapPlugin As SAP2000v20.cPluginCallback

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
        'default unit is meter
        SapModel.SetPresentUnits(Unitcomb.SelectedIndex + 1)  'SelectedIndex 從0開始

        Dim Point1, Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        Dim SelObj As New List(Of Member)

        Dim m1 As New Member

        Dim Zmax As Double = -999999
        Dim Zmin As Double = 999999

        If NumberItems < 2 Then
            MsgBox("At least 2 members need to be selected")
            GoTo 999
        End If


        For i = 0 To NumberItems - 1
            ret = SapModel.FrameObj.GetPoints(ObjectName(i), Point1, Point2)
            ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
            ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)


            '=====紀錄Zmax Zmin=========
            If Zmax < Math.Max(Z1, Z2) Then
                Zmax = Math.Max(Z1, Z2)
            End If

            If Zmin > Math.Min(Z1, Z2) Then
                Zmin = Math.Min(Z1, Z2)
            End If
            '===========================


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

            If Vector3D.CrossProduct(Vector1, Vector2).Length > 0.00001 Then
                MsgBox("選的桿件 " & SelObj(i).Name & " 和 " & SelObj(0).Name & " 非平行!")
                GoTo 999
            End If
        Next

        If Zmax - Zmin > 2 Then
            MsgBox("Selected members elevation range more then 2 meters, please re-select members which at the same elevation")
            GoTo 999
        End If











999:
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub PR_PointLoad_Import_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ret As Long
        Dim NumberNames As Long
        Dim MyName() As String

        ret = SapModel.LoadPatterns.GetNameList(NumberNames, MyName)

        For i = 0 To NumberNames - 1
            LoadPatterncomb.Items.Add(MyName(i))
        Next

        'Set Units
        LoadPatterncomb.SelectedIndex = 0  '第一個pattern
        Unitcomb.SelectedIndex = 7   'kg,m,C
        ComboBox3.SelectedIndex = 3  'Gravity


        'readSelectionRecord()

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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub
End Class