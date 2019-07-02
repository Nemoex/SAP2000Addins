Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports SAP2000v20
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Media.Media3D
Imports System.Runtime.InteropServices

Public Class Steel_Ratio_List

    Public SapModel As SAP2000v20.cSapModel
    Public ISapPlugin As SAP2000v20.cPluginCallback
    Public SectProp_Area As New Dictionary(Of String, Double)
    Public Points_Position As New Dictionary(Of String, Points_XYZ)
    Public Section_material As New Dictionary(Of String, String)
    Public Material_weight As New Dictionary(Of String, Double)
    Public GroupDict As New Dictionary(Of String, String)


    Public Structure Points_XYZ
        Dim X, Y, Z As Double
    End Structure

    'Private Declare Function SetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Long) As Long
    'Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
    'Private Const EM_REPLACESEL = &HC2
    'Private CharPos As Long
    'Private SendString As String
    'Dim EditHwnd As Long

    '    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    'Private Shared Function EnumChildWindows(ByVal WindowHandle As IntPtr, ByVal Callback As EnumWindowProcess, ByVal lParam As IntPtr) As Boolean
    '    End Function
    '    Public Delegate Function EnumWindowProcess(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
    '    Private Shared Function EnumWindow(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
    '        Dim ChildrenList As List(Of IntPtr) = GCHandle.FromIntPtr(Parameter).Target
    '        If ChildrenList Is Nothing Then Throw New Exception("GCHandle Target could not be cast as List(Of IntPtr)")
    '        ChildrenList.Add(Handle)
    '        Return True
    '    End Function

    Private Sub Steel_Ratio_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        RenewList()
        DataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect
        Wt_Calculate()

    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click

        Dim Y As Integer = DataGridView1.CurrentCellAddress.Y

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

        ret = SapModel.SelectObj.ClearSelection
        If Y < 0 Then GoTo 999
        ret = SapModel.FrameObj.SetSelected(DataGridView1.Item(1, Y).Value, True)

        'If ret = 0 Then
        '    MsgBox("    Frame  :  " & DataGridView1.Item(0, Y).Value & "   selected", MsgBoxStyle.SystemModal, "Select Frame")
        'End If
999:
        ret = SapModel.View.RefreshView


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If IsNumeric(TextBox1.Text) = False Or IsNumeric(TextBox2.Text) = False Then
            MsgBox("Please Check Steel Ratio Limit ", MsgBoxStyle.SystemModal, "Input Ratio Limit")
            GoTo 999
        End If
        If IsNumeric(TextBox1.Text) = True And IsNumeric(TextBox2.Text) = True Then
            If TextBox1.Text < TextBox2.Text Then
                MsgBox("Upper Limit < Over Limit !", MsgBoxStyle.SystemModal, "Check Steel Ratio Limit")
                GoTo 999
            End If
        End If



        RenewList()
        Wt_Calculate()

999:
    End Sub

    Private Function Wt_Calculate(Optional ByRef Wt1 As Double = 0, Optional ByRef Wt2 As Double = 0) As Boolean

        If DataGridView1.Item(1, 0).Value = "" Then
            MsgBox("Data Grid is Empty", MsgBoxStyle.SystemModal, "Steel Ratio List")
            GoTo 999
        End If

        Dim SteeltotalWeight1 As Double = 0
        Dim SteeltotalWeight2 As Double = 0
        Dim TotalWeightDiff As Double

        Dim WeightXRatio As Double = 0
        Dim AvgSteelRatio As Double = 0
        Dim Ratio As Double

        For i = 0 To DataGridView1.Rows.Count - 1
            SteeltotalWeight1 = SteeltotalWeight1 + DataGridView1.Item(7, i).Value * DataGridView1.Item(8, i).Value * DataGridView1.Item(9, i).Value
            SteeltotalWeight2 = SteeltotalWeight2 + DataGridView1.Item(10, i).Value * DataGridView1.Item(11, i).Value * DataGridView1.Item(12, i).Value
            If IsNumeric(DataGridView1.Item(1, i).Value) Then        'if ratio available
                'summary of ratio * design section weight
                If DataGridView1.Item(2, i).Value = "N/A" Then
                    Ratio = 0.0
                Else
                    Ratio = DataGridView1.Item(2, i).Value
                End If
                WeightXRatio = WeightXRatio + CDbl(Ratio) * DataGridView1.Item(10, i).Value * DataGridView1.Item(11, i).Value * DataGridView1.Item(12, i).Value
            End If
        Next
        'calculate Avg. Steel ratio        SUM(ratio * weight) / total weight
        AvgSteelRatio = WeightXRatio / SteeltotalWeight2


        Dim WeightBudget As Double
        If IsNumeric(TextBox3.Text) Then
            WeightBudget = TextBox3.Text
        ElseIf TextBox3.Text = "" Then
            WeightBudget = 0
        End If

        If CheckBox1.Checked Then
            If SteeltotalWeight2 - WeightBudget < 0 Then
                TextBox4.Text = "+ " + Format(WeightBudget - SteeltotalWeight2, "F3").ToString
                TextBox4.ForeColor = System.Drawing.Color.Black

            Else
                TextBox4.Text = "- " + Format(SteeltotalWeight2 - WeightBudget, "F3").ToString
                TextBox4.ForeColor = System.Drawing.Color.Red
            End If
        Else
            TextBox4.Text = "-----"
            TextBox4.ForeColor = System.Drawing.Color.Black
        End If

        TextBox5.Text = Format(SteeltotalWeight2, "F3")
        '先不顯示
        'Label12.Text = Format(AvgSteelRatio, "F3").ToString

        Wt1 = SteeltotalWeight1
        Wt2 = SteeltotalWeight2
        Wt_Calculate = True
999:

    End Function


    Private Sub RenewList()

        Points_Position.Clear()
        Section_material.Clear()
        Material_weight.Clear()
        GroupDict.Clear()

        Dim ListExist As Boolean
        Dim ListSequency() As String
        If DataGridView1.RowCount > 1 Then
            ListExist = True
            ReDim ListSequency(DataGridView1.RowCount - 1)
            For i = 0 To DataGridView1.RowCount - 1
                ListSequency(i) = DataGridView1.Item(1, i).Value
                If DataGridView1.Item(1, i).Value <> Nothing Then
                    GroupDict.Add(DataGridView1.Item(1, i).Value, "")
                End If
            Next
            getGroupInfo()
        End If

        '====create GroupDictionary
        'For i = 0 To DataGridView1.RowCount - 1
        '    GroupDict.Add(DataGridView1.Item(1, i).Value, "_")
        'Next
        'getGroupInfo()
        '==========================

        Dim RowIndex As Integer
        RowIndex = DataGridView1.FirstDisplayedScrollingRowIndex

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

        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.PointObj.SetSelected("ALL", True, eItemType.Group)
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        Dim Pt_Pos As New Points_XYZ

        For i = 0 To NumberItems - 1
            ret = SapModel.PointObj.GetCoordCartesian(ObjectName(i), Pt_Pos.X, Pt_Pos.Y, Pt_Pos.Z)
            Points_Position.Add(ObjectName(i), Pt_Pos)
        Next
        '=======================

        'get section material
        Dim NameInFile As String
        Dim FileName As String
        Dim MatProp As String
        Dim PropType As eFramePropType


        Dim NumberNames As Long
        Dim MyName() As String

        Dim w As Double
        Dim m As Double
        ret = SapModel.PropFrame.GetNameList(NumberNames, MyName)

        For i = 0 To NumberNames - 1
            ret = SapModel.PropFrame.GetNameInPropFile(MyName(i), NameInFile, FileName, MatProp, PropType)
            ret = SapModel.PropMaterial.GetWeightAndMass(MatProp, w, m)

            If MatProp Is Nothing Then

                MsgBox("Can't Get " & MyName(i) & "Material Data Plaese Check!")
                GoTo FSEC
            End If

            Section_material.Add(MyName(i), MatProp)


            If Material_weight.ContainsKey(MatProp) Then

            Else
                Material_weight.Add(MatProp, w)
            End If
FSEC:
        Next







        '=======================



        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.All
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        'If ListExist = True Then
        '    ObjectName = listsequency
        'End If

        Dim Area1, Area2 As Double
        Dim L1, L2 As Double
        Dim Wt1, Wt2 As Double
        Dim KLROver As Boolean
        Dim SlenderRatio As String
        Dim PhysicalGroup As String

        SapModel.SetPresentUnits(eUnits.Ton_m_C)

        If ListExist = True Then
            For i = 0 To ListSequency.Count - 2
                SlenderRatio = ""
                ret = SapModel.DesignSteel.GetSummaryResults(ListSequency(i), NumberResults, FrameName, Ratio, RatioType, Location, ComboName, ErrorSummary, WarningSummary)
                If ErrorSummary IsNot Nothing And WarningSummary IsNot Nothing And NumberResults > 0 Then
                    If ErrorSummary(0).Contains("kl/r") Or WarningSummary(0).Contains("kl/r") Then
                        SlenderRatio = "Over"
                    End If
                End If

                'If NumberResults = 0 Then GoTo ignoreRC
                ret = SapModel.FrameObj.GetSection(ListSequency(i), BeginSection, SAuto)
                ret = SapModel.DesignSteel.GetDesignSection(ListSequency(i), DesignSection)

                'If Ratio Is Nothing Then
                'If NumberResults = 0 Then
                'ReDim Ratio(0)
                'Ratio(0) = "0.0"
                'Ratio(0) = "N/A"
                'End If
                If ComboName Is Nothing Then
                    ReDim ComboName(0)
                    ComboName(0) = ""
                End If

                getFrameData(ListSequency(i), BeginSection, Area1, L1, Wt1)
                getFrameData(ListSequency(i), DesignSection, Area2, L2, Wt2)



                PhysicalGroup = getPhyGroup(ListSequency(i))

                Dim row As String()
                'If KLROver = True Then
                If NumberResults = 0 Then
                    row = New String() {PhysicalGroup, CInt(ListSequency(i)), "N/A", SlenderRatio, ComboName(0), BeginSection, DesignSection, Area1, L1, Wt1, Area2, L2, Wt2}
                Else
                    row = New String() {PhysicalGroup, CInt(ListSequency(i)), Format(Ratio(0), "F4"), SlenderRatio, ComboName(0), BeginSection, DesignSection, Area1, L1, Wt1, Area2, L2, Wt2}
                End If
                'row = New String() {CInt(ListSequency(i)), Format(Ratio(0), "F4"), SlenderRatio, ComboName(0), BeginSection, DesignSection, Area1, L1, Wt1, Area2, L2, Wt2}
                'Else
                'row = New String() {CInt(ObjectName(i)), Format(Ratio(0), "F4"), "", ComboName(0), BeginSection, DesignSection, Area1, L1, Wt1, Area2, L2, Wt2}
                'End If
                '=============

                DataGridView1.Rows.Add(row)
                If BeginSection <> DesignSection Then
                    DataGridView1.Item(6, memberCounter).Style.ForeColor = System.Drawing.Color.Red
                End If

                memberCounter += 1
            Next
            GoTo ListFin
        End If

        '====create GroupDictionary
        For i = 0 To ObjectType.Count - 1
            If ObjectType(i) = 2 Then
                GroupDict.Add(ObjectName(i), "")
            End If
        Next
        getGroupInfo()
        '==========================



        For i = 0 To ObjectType.Count - 1
            If ObjectType(i) = 2 Then    '2 = Frame object

                'KLROver = False
                SlenderRatio = ""
                ret = SapModel.DesignSteel.GetSummaryResults(ObjectName(i), NumberResults, FrameName, Ratio, RatioType, Location, ComboName, ErrorSummary, WarningSummary)
                If ErrorSummary IsNot Nothing And WarningSummary IsNot Nothing Then
                    If ErrorSummary(0).Contains("kl/r") Or WarningSummary(0).Contains("kl/r") Then
                        'KLROver = True
                        SlenderRatio = "Over"
                    End If
                End If

                If NumberResults = 0 Then GoTo NoSteelResult
                ret = SapModel.FrameObj.GetSection(ObjectName(i), BeginSection, SAuto)
                ret = SapModel.DesignSteel.GetDesignSection(ObjectName(i), DesignSection)

                If Ratio Is Nothing Then
                    ReDim Ratio(0)
                    Ratio(0) = 0.0
                End If
                If ComboName Is Nothing Then
                    ReDim ComboName(0)
                    ComboName(0) = ""
                End If

                getFrameData(ObjectName(i), BeginSection, Area1, L1, Wt1)
                getFrameData(ObjectName(i), DesignSection, Area2, L2, Wt2)

                PhysicalGroup = getPhyGroup(ObjectName(i))
                'Dim row As String() = New String() {ObjectName(i), Format(Ratio(0), "F3"), ComboName(0), BeginSection, DesignSection, Area1, L1, Wt1, Area2, L2, Wt2}
                '=============kl/r core
                Dim row As String()
                'If KLROver = True Then

                row = New String() {PhysicalGroup, CInt(ObjectName(i)), Format(Ratio(0), "F4"), SlenderRatio, ComboName(0), BeginSection, DesignSection, Area1, L1, Wt1, Area2, L2, Wt2}
                'Else
                'row = New String() {CInt(ObjectName(i)), Format(Ratio(0), "F4"), "", ComboName(0), BeginSection, DesignSection, Area1, L1, Wt1, Area2, L2, Wt2}
                'End If
                '=============

                DataGridView1.Rows.Add(row)
                If BeginSection <> DesignSection Then
                    DataGridView1.Item(6, memberCounter).Style.ForeColor = System.Drawing.Color.Red
                End If

                memberCounter += 1
NoSteelResult:
            End If
        Next
ListFin:
        DataGridView1.Columns(0).ValueType = GetType(Integer)

        ret = SapModel.SelectObj.ClearSelection

        If ListExist = False Then
            DataGridView1.Sort(DataGridView1.Columns(2), System.ComponentModel.ListSortDirection.Ascending)
        End If

        DataGridView1.FirstDisplayedScrollingRowIndex = RowIndex

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        RenewList()

        Dim Wt1, Wt2 As Double
        Wt_Calculate(Wt1, Wt2)

        If IsNumeric(TextBox3.Text) = False Then
            MsgBox("Please Check Weight Budget Value", MsgBoxStyle.SystemModal, "Check Input Value")
            GoTo 999
        End If

        If SapModel.GetModelIsLocked = False Then
            GoTo 999
        End If

        Dim TotalWeightDiff As Double
        TotalWeightDiff = Wt2 - Wt1

        If TotalWeightDiff = 0 Then
            MsgBox("There is No Need To OverWrite Section", MsgBoxStyle.SystemModal, "Overwrite Section Type")
            GoTo 999
        End If

        Dim Message As String
        If TotalWeightDiff > 0 Then
            Message = "Increase  "
        Else
            Message = "Decrease  "
        End If

        If TotalWeightDiff <> 0 Then
            MsgBox("Total Weight (Ton)          :     " & vbCrLf & Format(Wt1, "F3") & "  >>>  " & Format(Wt2, "F3") & vbCrLf & Message & Format(Math.Abs(TotalWeightDiff), "F3") & "  Ton", MsgBoxStyle.SystemModal, "Steel Weight Calculate")
        End If

        If MsgBox("Unlock Model and Overwrite Section Change", MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal, "Overwrite Section Change") = MsgBoxResult.Yes Then

            Dim ret As Long
            Dim NumberItems As Long
            Dim ObjectType() As Integer
            Dim ObjectName() As String
            Dim memberCounter As Integer
            Dim PropName As String
            Dim BeginSection As String
            Dim SAuto As String

            ret = SapModel.SetModelIsLocked(False)

            ret = SapModel.SelectObj.All
            ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
            For i = 0 To ObjectType.Count - 1
                If ObjectType(i) = 2 Then    '2 = Frame object
                    ret = SapModel.FrameObj.GetSection(ObjectName(i), BeginSection, SAuto)

                    ret = SapModel.DesignSteel.GetDesignSection(ObjectName(i), PropName)
                    If BeginSection <> PropName Then
                        ret = SapModel.FrameObj.SetSection(ObjectName(i), PropName)
                        memberCounter += 1
                    End If
                End If
            Next

            If memberCounter >= 1 Then
                MsgBox("Overwrite Section Number    :    " & memberCounter & vbCrLf & "Please Analysize Model Again", MsgBoxStyle.SystemModal, "Overwrite Complete")

            End If
999:

        End If
    End Sub

    Private Function getFrameData(ByVal FrameNum As String, ByVal SectionName As String, ByRef Area1 As Double, ByRef Length As Double, ByRef U_Wt As Double) As Boolean
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



        If SectProp_Area.ContainsKey(SectionName) Then
            Area = SectProp_Area.Item(SectionName)
        Else
            ret = SapModel.PropFrame.GetSectProps(SectionName, Area, as2, as3, Torsion, I22, I33, S22, S33, Z22, Z33, R22, R33)
            SectProp_Area.Add(SectionName, Area)
        End If

        Area1 = Area
        Dim Point1, Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        ret = SapModel.FrameObj.GetPoints(FrameNum, Point1, Point2)

        Dim Pt1, Pt2 As New Points_XYZ

        Pt1 = Points_Position.Item(Point1)
        Pt2 = Points_Position.Item(Point2)

        'ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
        'ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)

        Length = ((Pt1.X - Pt2.X) ^ 2 + (Pt1.Y - Pt2.Y) ^ 2 + (Pt1.Z - Pt2.Z) ^ 2) ^ 0.5

        '==================
        Dim MatProp As String

        MatProp = Section_material(SectionName)

        '==================
        'Dim NameInFile As String
        'Dim FileName As String
        'Dim MatProp As String
        'Dim PropType As eFramePropType

        'Dim t3 As Double
        'Dim t2 As Double
        'Dim tf As Double
        'Dim tw As Double
        'Dim t2b As Double
        'Dim tfb As Double
        'Dim Color As Long
        'Dim Notes As String
        'Dim GUID As String
        'Dim dis As Double
        'Dim Thickness As Double
        'Dim Radius As Double
        'Dim LipDepth As Double
        'Dim LipAngle As Double

        'Dim NumberItems As Long
        'Dim ShapeName() As String
        'Dim MyType() As Long
        'Dim DesignType As Long

        'ret = SapModel.PropFrame.GetNameInPropFile(SectionName, NameInFile, FileName, MatProp, PropType)
        'If MatProp Is Nothing Then
        '    ret = SapModel.PropFrame.GetType(SectionName, PropType)
        '    Select Case PropType
        '        Case eFramePropType.SECTION_I
        '            ret = SapModel.PropFrame.GetISection(SectionName, FileName, MatProp, t3, t2, tf, tw, t2b, tfb, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_CHANNEL
        '            ret = SapModel.PropFrame.GetChannel(SectionName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_T
        '            ret = SapModel.PropFrame.GetTee(SectionName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_ANGLE
        '            ret = SapModel.PropFrame.GetAngle(SectionName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_DBLANGLE
        '            ret = SapModel.PropFrame.GetDblAngle(SectionName, FileName, MatProp, t3, t2, tf, tw, dis, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_BOX

        '        Case eFramePropType.SECTION_PIPE
        '            ret = SapModel.PropFrame.GetPipe(SectionName, FileName, MatProp, t3, tw, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_RECTANGULAR
        '            ret = SapModel.PropFrame.GetRectangle(SectionName, FileName, MatProp, t3, t2, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_CIRCLE
        '            ret = SapModel.PropFrame.GetCircle(SectionName, FileName, MatProp, t3, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_GENERAL
        '            ret = SapModel.PropFrame.GetGeneral(SectionName, FileName, MatProp, t3, t2, Area, as2, as3, Torsion, I22, I33, S22, S33, Z22, Z33, R22, R33, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_DBCHANNEL
        '            ret = SapModel.PropFrame.GetDblChannel(SectionName, FileName, MatProp, t3, t2, tf, tw, dis, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_AUTO

        '        Case eFramePropType.SECTION_SD
        '            ret = SapModel.PropFrame.GetSDSection(SectionName, MatProp, NumberItems, ShapeName, MyType, DesignType, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_VARIABLE

        '        Case eFramePropType.SECTION_JOIST

        '        Case eFramePropType.SECTION_BRIDGE

        '        Case eFramePropType.SECTION_COLD_C
        '            ret = SapModel.PropFrame.GetColdC(SectionName, FileName, MatProp, t3, t2, Thickness, Radius, LipDepth, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_COLD_2C

        '        Case eFramePropType.SECTION_COLD_Z
        '            ret = SapModel.PropFrame.GetColdZ(SectionName, FileName, MatProp, t3, t2, Thickness, Radius, LipDepth, LipAngle, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_COLD_L

        '        Case eFramePropType.SECTION_COLD_2L

        '        Case eFramePropType.SECTION_COLD_HAT
        '            ret = SapModel.PropFrame.GetColdHat(SectionName, FileName, MatProp, t3, t2, Thickness, Radius, LipDepth, Color, Notes, GUID)
        '        Case eFramePropType.SECTION_BUILTUP_I_COVERPLATE

        '        Case eFramePropType.SECTION_PCC_GIRDER_I
        '            'RC
        '        Case eFramePropType.SECTION_PCC_GIRDER_U
        '            'RC
        '    End Select

        'End If

        'Dim w As Double
        'Dim m As Double
        'ret = SapModel.PropMaterial.GetWeightAndMass(MatProp, w, m)
        'U_Wt = w
        U_Wt = Material_weight.Item(MatProp)
        If U_Wt <> 0 Then Return True

    End Function

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If e.ColumnIndex = 2 Then

            If e.Value = "N/A" Then
                DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = System.Drawing.Color.Red
                GoTo NoSteelResult
            End If

            If IsNumeric(TextBox1.Text) = True Then
                If e.Value > CDbl(TextBox1.Text) Then
                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = System.Drawing.Color.Red
                End If
                If e.Value >= CDbl(TextBox2.Text) And e.Value <= CDbl(TextBox1.Text) Then
                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = System.Drawing.Color.Black
                End If
            End If
            If IsNumeric(TextBox2.Text) = True Then
                If e.Value < CDbl(TextBox2.Text) Then
                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = System.Drawing.Color.Blue
                End If
                If e.Value >= CDbl(TextBox2.Text) And e.Value <= CDbl(TextBox1.Text) Then
                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = System.Drawing.Color.Black
                End If
            End If
        End If
NoSteelResult:
    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        DataGridView1.ClearSelection()
        DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Selected = True
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        DataGridView1.ClearSelection()
        DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Selected = True
    End Sub

    Private Sub DataGridView1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDown
        DataGridView1.ClearSelection()
        DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Selected = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.Column6.Visible = False Then
            Me.Column6.Visible = True
            Me.Column7.Visible = True
            Me.Column8.Visible = True
            Me.Column9.Visible = True
            Me.Column10.Visible = True
            Me.Column11.Visible = True
        Else
            Me.Column6.Visible = False
            Me.Column7.Visible = False
            Me.Column8.Visible = False
            Me.Column9.Visible = False
            Me.Column10.Visible = False
            Me.Column11.Visible = False
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked Then
            TextBox3.ReadOnly = False
            Wt_Calculate()
            Label7.Visible = True
            Label8.Visible = True
        Else
            TextBox3.ReadOnly = True
            TextBox4.ReadOnly = True
            TextBox4.Text = "-----"
            TextBox4.ForeColor = System.Drawing.Color.Black
            Label7.Visible = False
            Label8.Visible = False
        End If

    End Sub

    Private Sub getFrameGroup()



        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        Dim counter As Integer

        Dim ret As Long
        Dim MyName() As String
        Dim NumberNames As Long
        Dim Name As String
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim Y As Integer = DataGridView1.CurrentCellAddress.Y

        ret = SapModel.GroupDef.GetNameList(NumberNames, MyName)
        For i = 0 To NumberNames - 1
            ret = SapModel.GroupDef.GetAssignments(MyName(i), NumberItems, ObjectType, ObjectName)
            For j = 0 To NumberItems - 1
                If ObjectType(j) = 2 And ObjectName(j) = DataGridView1.Item(1, Y).Value Then
                    ComboBox1.Items.Add(MyName(i))
                End If
            Next
        Next
        ret = SapModel.SelectObj.ClearSelection


        If ComboBox1.Items.Count > 0 Then
            ComboBox1.Text = "Select Group"
        Else
            ComboBox1.Text = ""
        End If


        ComboBox1.Visible = True
        ComboBox1.SelectedIndex = -1

        'get section list
        ret = SapModel.PropFrame.GetNameList(NumberNames, MyName)

        For i = 0 To NumberNames - 1
            ComboBox2.Items.Add(MyName(i))
        Next

        ComboBox2.Visible = True
        ComboBox2.Text = ""
        ComboBox2.SelectedText = DataGridView1.Item(5, Y).Value

    End Sub

    Private Sub getGroupInfo()
        Dim ret As Long
        Dim MyName() As String
        Dim NumberNames As Long
        Dim Name As String
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim GroupCount As Integer = 0
        Dim counter As Integer

        ret = SapModel.GroupDef.GetNameList(NumberNames, MyName)
        For i = 0 To NumberNames - 1
            'get each group components and build group information dictionary
            ret = SapModel.GroupDef.GetAssignments(MyName(i), NumberItems, ObjectType, ObjectName)
            For j = 0 To NumberItems - 1
                If ObjectType(j) = 2 Then
                    GroupDict(ObjectName(j)) = GroupDict(ObjectName(j)) + MyName(i) + ","
                End If
            Next
        Next

    End Sub

    Private Function getPhyGroup(ByRef FrameNumber As String) As String
        Dim GroupString As String
        GroupString = GroupDict(FrameNumber)
        Dim Groups() As String
        Dim charseparators() As Char = ","
        Groups = GroupString.Split(charseparators, StringSplitOptions.RemoveEmptyEntries)
        Dim GroupCount As Integer = 0

        For i = 0 To Groups.Count - 1
            If IsNumeric(Groups(i)) Then
                getPhyGroup = Groups(i)
                GroupCount += 1
                GoTo fin
            Else
                GroupCount += 1
            End If
        Next
        getPhyGroup = "**"
fin:
        If GroupCount = 0 Then
            getPhyGroup = "None"
        End If
        Return getPhyGroup

    End Function


    Private Function getPhysicalGroup(ByVal FrameName As String) As String

        Dim ret As Long
        Dim MyName() As String
        Dim NumberNames As Long
        Dim Name As String
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String

        Dim GroupCount As Integer = 0

        Dim counter As Integer

        'get group list
        ret = SapModel.GroupDef.GetNameList(NumberNames, MyName)
        For i = 0 To NumberNames - 1
            'get each group components
            ret = SapModel.GroupDef.GetAssignments(MyName(i), NumberItems, ObjectType, ObjectName)
            For j = 0 To NumberItems - 1
                If ObjectType(j) = 2 And ObjectName(j) = FrameName Then
                    If IsNumeric(MyName(i)) Then
                        getPhysicalGroup = MyName(i)
                        GroupCount += 1
                        GoTo exitFunc
                    Else
                        GroupCount += 1
                    End If
                End If
            Next
        Next
        getPhysicalGroup = "**"
exitFunc:

        If GroupCount = 0 Then
            getPhysicalGroup = "None"
            'ElseIf GroupCount = 1 Then
            '    'no need change data
            'Else
            '    getPhysicalGroup = "*" + getPhysicalGroup
        End If

        Return getPhysicalGroup

    End Function



    'Public Shared Function GetChildWindows(ByVal ParentHandle As IntPtr) As IntPtr()
    '    Dim ChildrenList As New List(Of IntPtr)
    '    Dim ListHandle As GCHandle = GCHandle.Alloc(ChildrenList)
    '    Try
    '        EnumChildWindows(ParentHandle, AddressOf EnumWindow, GCHandle.ToIntPtr(ListHandle))
    '    Finally
    '        If ListHandle.IsAllocated Then ListHandle.Free()
    '    End Try

    '    Return ChildrenList.ToArray
    'End Function

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Space Then
            If IsNumeric(TextBox3.Text) Then
                Wt_Calculate()
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim ret As Long
        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.Group(ComboBox1.SelectedItem.ToString)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim Y As Integer = DataGridView1.CurrentCellAddress.Y
        Dim ret As Long

        If ComboBox1.SelectedIndex > 0 Then
            ret = SapModel.DesignSteel.SetDesignSection(ComboBox1.Text, ComboBox2.Text, False, eItemType.Group)
        Else
            ret = SapModel.DesignSteel.SetDesignSection(DataGridView1.Item(1, Y).Value, ComboBox2.Text, False, eItemType.Objects)
        End If

        RenewList()
        ret = SapModel.View.RefreshView

    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        getFrameGroup()
    End Sub

    Private Sub DataGridView1_SortCompare(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewSortCompareEventArgs) Handles DataGridView1.SortCompare
        If e.Column.Index = 1 Then
            e.SortResult = If(CInt(e.CellValue1) < CInt(e.CellValue2), -1, 1)
            e.Handled = True
        End If
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        If CheckBox1.Checked Then
            CheckBox1.Checked = False
        Else
            CheckBox1.Checked = True
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If DataGridView1.SortOrder = SortOrder.None Then
            DataGridView1.Sort(DataGridView1.Columns(2), System.ComponentModel.ListSortDirection.Ascending)
        ElseIf DataGridView1.SortOrder = SortOrder.Ascending Then
            DataGridView1.Sort(DataGridView1.Columns(2), System.ComponentModel.ListSortDirection.Descending)
        ElseIf DataGridView1.SortOrder = SortOrder.Descending Then
            DataGridView1.Sort(DataGridView1.Columns(2), System.ComponentModel.ListSortDirection.Ascending)
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If DataGridView1.RowCount <= 1 Then GoTo endsub2

        'SaveFileDialog1.Title = "Save Data"

        'Dim misValue As Object = System.Reflection.Missing.Value

        ''SaveFileDialog1.Filter = "xlsx"
        'If SaveFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
        '    'FileOpen(10, SaveFileDialog1.FileName, OpenMode.Output)
        '    '#執行一個新的Excel Application        
        '    xlApp = CreateObject("Excel.Application")
        '    xlBook = xlapp.Workbooks.Open(SaveFileDialog1.FileName)
        '    'FileClose(10)
        'Else
        '    GoTo endsub2
        'End If
        'Call Excel API

        On Error Resume Next
        xlApp = CreateObject("Excel.Application")
        xlBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value)


        '停用警告訊息        
        'xlApp.DisplayAlerts = False
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

        'For i = 0 To DataGridView1.RowCount - 1
        '    For j = 0 To DataGridView1.ColumnCount - 1
        '        For k = 1 To DataGridView1.ColumnCount
        '            xlSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
        '            xlSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString
        '            count += 1
        '        Next
        '    Next
        'Next

        MsgBox("Finish", MsgBoxStyle.SystemModal, "Output Excel")
        'xlBook.SaveAs(SaveFileDialog1.FileName)
        'xlBook.Close()

        'System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
        'xlApp = Nothing
        'xlBook = Nothing
        'xlSheet = Nothing
        'xlRange = Nothing
        'GC.Collect()

endsub2:
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        If Me.Column6.Visible = False Then
            Me.Column6.Visible = True
            Me.Column7.Visible = True
            Me.Column8.Visible = True
            Me.Column9.Visible = True
            Me.Column10.Visible = True
            Me.Column11.Visible = True
            Me.Column5.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Else
            Me.Column6.Visible = False
            Me.Column7.Visible = False
            Me.Column8.Visible = False
            Me.Column9.Visible = False
            Me.Column10.Visible = False
            Me.Column11.Visible = False
            Me.Column5.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If
    End Sub

    Private Sub ExpandableSplitter1_ExpandedChanged(ByVal sender As System.Object, ByVal e As DevComponents.DotNetBar.ExpandedChangeEventArgs) Handles ExpandableSplitter1.ExpandedChanged
        If ExpandableSplitter1.Expanded = True Then
            DataGridView1.Size = New System.Drawing.Size(567, 392)
        Else
            DataGridView1.Size = New System.Drawing.Size(715, 392)
        End If
    End Sub


    Private Sub SwitchButton1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SwitchButton1.ValueChanged
        Dim RatioXWt As Double
        Dim SteelRatio As Double
        Dim SteeltotalWeight As Double

        If DataGridView1.Item(1, 0).Value = "" Then GoTo EndSub

        For i = 0 To DataGridView1.RowCount - 1
            DataGridView1.Item(13, i).Value = 0
        Next



        If SwitchButton1.Value = True Then
            For i = 0 To DataGridView1.Rows.Count - 1
                If IsNumeric(DataGridView1.Item(2, i).Value) Then
                    SteelRatio = DataGridView1.Item(2, i).Value
                ElseIf DataGridView1.Item(2, i).Value = "N/A" Then
                    SteelRatio = 0.0
                End If
                SteeltotalWeight = SteeltotalWeight + DataGridView1.Item(10, i).Value * DataGridView1.Item(11, i).Value * DataGridView1.Item(12, i).Value
                RatioXWt = RatioXWt + SteelRatio * DataGridView1.Item(10, i).Value * DataGridView1.Item(11, i).Value * DataGridView1.Item(12, i).Value
            Next
        Else
            '====暫時關閉=========================
            For i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Item(0, i).Value = "None" Or DataGridView1.Item(0, i).Value = "**" Then
                    DataGridView1.Item(13, i).Value = DataGridView1.Item(2, i).Value
                    GoTo endsearch
                End If
                For j = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Item(0, j).Value = DataGridView1.Item(0, i).Value And DataGridView1.Item(0, j).Value <> "None" And DataGridView1.Item(0, j).Value <> "**" Then
                        If DataGridView1.Item(2, j).Value > DataGridView1.Item(13, i).Value Then
                            DataGridView1.Item(13, i).Value = DataGridView1.Item(2, j).Value
                        End If
                    End If
                Next
endsearch:
            Next
            '=====================================
            For i = 0 To DataGridView1.Rows.Count - 1
                If IsNumeric(DataGridView1.Item(13, i).Value) Then
                    SteelRatio = DataGridView1.Item(13, i).Value
                ElseIf DataGridView1.Item(13, i).Value = "N/A" Then
                    SteelRatio = 0.0
                End If
                SteeltotalWeight = SteeltotalWeight + DataGridView1.Item(10, i).Value * DataGridView1.Item(11, i).Value * DataGridView1.Item(12, i).Value
                RatioXWt = RatioXWt + SteelRatio * DataGridView1.Item(10, i).Value * DataGridView1.Item(11, i).Value * DataGridView1.Item(12, i).Value
            Next
        End If

        Label12.Text = Format(RatioXWt / SteeltotalWeight, "F3")

EndSub:
    End Sub






End Class