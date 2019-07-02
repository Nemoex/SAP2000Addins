Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports Sunisoft.IrisSkin
Imports SAP2000v20





Public Class Menu

    Public SapModel As SAP2000v20.cSapModel
    Public ISapPlugin As SAP2000v20.cPluginCallback
    Public elev() As Double
    Public Lcount As Integer
    Public ConnectiveJoint() As Integer
    Public ConnCount As Integer
    Public GroupDict As New Dictionary(Of String, String)
    Public ProgramStartCnt As New ProgStartCnt.ProgCnt


    Private Sub Menu_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        SE.Active = False
        frmMenu = Nothing
        ISapPlugin.Finish(0)



        Try
            If My.Computer.FileSystem.FileExists("C:\SAPaddinTemp1.ini") Then
                My.Computer.FileSystem.DeleteFile("C:\SAPaddinTemp1.ini")
            End If
            If My.Computer.FileSystem.FileExists("C:\SAPaddinTemp2.ini") Then
                My.Computer.FileSystem.DeleteFile("C:\SAPaddinTemp2.ini")
            End If
        Catch ex As Exception
            MsgBox("Warning : Permission")
        End Try



    End Sub


    Private Sub Menu_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If Asc(e.KeyChar) = 21 Then
            If ComboBox1.Visible = False Then
                ComboBox1.Visible = True
                Button16.Visible = True
                TextBox1.Visible = True
            Else
                ComboBox1.Visible = False
                Button16.Visible = False
                TextBox1.Visible = False
            End If
        End If
    End Sub


    'CTCI環境檢查
    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Dim checkCTCIGroup As Boolean = False

        'If System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "ctci.com.tw" Then
        '    checkCTCIGroup = True
        'ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "jdec.com.cn" Then
        '    checkCTCIGroup = True
        'ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "cimas.com.vn" Then
        '    checkCTCIGroup = True
        'ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "cinda.in" Then
        '    checkCTCIGroup = True
        'ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "ctci.co.th" Then
        '    checkCTCIGroup = True
        'ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "ctcim.com.tw" Then
        '    checkCTCIGroup = True
        'End If

        If CheckDomain() = False Then
            MsgBox("This Plug-in is for CTCI Group Only")
            Me.Close()
        End If

        Button10.Text = "Classify RC" & vbCrLf & "Piece Mark"
        Button11.Text = "  Output RC " & vbCrLf & "  Design Result"
        Button12.Text = "Loading Input for" & vbCrLf & "Piperack(Point)"
        Button14.Text = "Loading Input for" & vbCrLf & "Piperack(Distributed)"
        Button17.Text = "Start End Rule" & vbCrLf & "Check"
        Button19.Text = "Wind Load Area" & vbCrLf & "Calculation"
        Button20.Text = "New Piping Load" & vbCrLf & "input tool"
        btnmodifyASD_CaseComb.Text = "modify nonlinear" & vbCrLf & "case/comb for ASD"

        SkinEngine1.Active = True
        'SkinEngine1.SerialNumber = ""




    End Sub

    '刪除所有"數字"的Group後重新建立
    Private Sub Create_Group(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ret As Long
        Dim NumberNames As Long
        Dim GroupName() As String


        If MsgBox("Unlock Model and Create Group Number", MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal, "Auto Create Group") = MsgBoxResult.Yes Then
            ret = SapModel.SetModelIsLocked(False)
        Else
            GoTo endSub
        End If

        ret = SapModel.GroupDef.GetNameList(NumberNames, GroupName)

        For i = 0 To NumberNames - 1
            If GroupName(i) <> "All" And IsNumeric(GroupName(i)) = True Then
                ret = SapModel.GroupDef.Delete(GroupName(i))
            End If
        Next

        ret = SapModel.SelectObj.All
        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        For i = 0 To NumberItems - 1
            If ObjectType(i) = 2 Then
                ret = SapModel.GroupDef.SetGroup(ObjectName(i))
                ret = SapModel.FrameObj.SetGroupAssign(ObjectName(i), ObjectName(i))
            End If
        Next
        MsgBox("Create Member Group Complete", , "Auto Create Group")
endSub:

        ret = SapModel.SelectObj.ClearSelection
        ISapPlugin.Finish(0)

        ProgramStartCnt.inputName("SAP-Create Group")

    End Sub

    '刪除"數字"群組
    Private Sub Delete_Group(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ret As Long
        Dim NumberNames As Long
        Dim GroupName() As String

        If MsgBox("Unlock Model and Delete All Group", MsgBoxStyle.YesNo, "Delete All Group") = MsgBoxResult.Yes Then
            ret = SapModel.SetModelIsLocked(False)
        Else
            GoTo endSub
        End If

        ret = SapModel.GroupDef.GetNameList(NumberNames, GroupName)

        For i = 0 To NumberNames - 1
            If GroupName(i).ToString.ToUpper <> "ALL" And IsNumeric(GroupName(i)) Then
                ret = SapModel.GroupDef.Delete(GroupName(i))
            End If
        Next

        MsgBox("Delete Member Group Complete", , "Delete All Numeric Group")
endSub:
        ISapPlugin.Finish(0)

        ProgramStartCnt.inputName("SAP-Delete Group")
    End Sub

    '撓度檢核
    Private Sub Deflection_Check(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ret As Long
        Dim NumberResults As Long
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


        Dim NumberCombs As Integer
        'Dim MyName2() As String
        'Dim Selected As Boolean
        Dim NumberItems As Integer
        Dim CombName() As String
        'Dim selectedCL As String

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
            MsgBox("Need Run Analysis First" & vbCrLf & "End Program", , "Deflection Check")
            GoTo ExitSub
        End If

        ret = SapModel.RespCombo.GetNameList(NumberCombs, CombName)


        ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
        'ret = SapModel.Results.Setup.SetCaseSelectedForOutput("COMB1")

        'selectedCL = InputBox("Load Case/Combination :")
        'selectedCL = selectedCL.ToUpper
        'ret = SapModel.Results.Setup.SetComboSelectedForOutput(selectedCL)

        'Dim SelectedFrame As String
        'Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        '=======20140715==============
        Dim NumberItems2 As Integer
        Dim ObjectType2() As Integer
        Dim ObjectName2() As String
        ret = SapModel.SelectObj.All
        ret = SapModel.SelectObj.GetSelected(NumberItems2, ObjectType2, ObjectName2)

        '清除與建立空的Group Dictionary
        GroupDict.Clear()
        For i = 0 To NumberItems2 - 1
            If ObjectType2(i) = 2 Then
                GroupDict.Add(ObjectName2(i), "")
            End If
        Next
        ret = SapModel.SelectObj.ClearSelection
        '=============================



        'get point displacements
        SapModel.SetPresentUnits(eUnits.Ton_mm_C)
        'ret = SapModel.PointObj.SetSelected("20", True, eItemType.Object)
        'ret = SapModel.Results.JointDispl("29", eItemTypeElm.GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)


        'ret = SapModel.PointObj.GetSpecialPoint("~85", True)

        'Dim NumberItems As Integer
        'Dim ObjectType() As Integer
        'Dim ObjectName() As String
        'Dim PointNumber() As Integer
        'ret = SapModel.PointObj.GetConnectivity("29", NumberItems, ObjectType, ObjectName, PointNumber)


        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double
        Dim X3, Y3, Z3 As Double
        Dim SDef, EDef, MDef As Double

        Dim SPt, EPt, MPt As Point3D
        Dim Result As Double
        Dim MaxResult As Double = 999999
        Dim MaxNodeFlag As Integer
        Dim MaxCombFlag As Integer

        'Dim SegmentCounter As Integer
        Dim VoidPointCounter As Integer

        Dim AutoMesh As Boolean
        Dim AutoMeshAtPoints As Boolean
        Dim AutoMeshAtLines As Boolean
        Dim NumSegs As Long
        Dim AutoMeshMaxLength As Double
        Dim AutoSelectFrame As MsgBoxResult
        Dim ParallelTo(5) As Boolean

        If NumberItems = 0 Then
            AutoSelectFrame = MsgBox("Didn't select any frame for check" & vbCrLf & "Do you want to select ""ALL Horizontal Frame"" to check ? ", MsgBoxStyle.YesNo, "Deflection Check")
            If AutoSelectFrame = MsgBoxResult.Yes Then
                ParallelTo(0) = True
                ParallelTo(1) = True
                ret = SapModel.SelectObj.LinesParallelToCoordAxis(ParallelTo)

                ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
                '=======20140715==============
                '清除與建立空的Group Dictionary
                'GroupDict.Clear()
                'For i = 0 To NumberItems - 1
                '    If ObjectType(i) = 2 Then
                '        GroupDict.Add(ObjectName(i), "")
                '    End If
                'Next
                '=============================

            Else
                MsgBox("Please select frame first then re-start command", , "Deflection Check")
                GoTo ExitSub
            End If
        End If

        '======20140715========
        '建立Frame Name 與 Group 之關係
        getGroupInfo()
        '======================


        Dim SelectCombFrm As New SelectCombDialog

        For i = 0 To NumberCombs - 1
            SelectCombFrm.ListBox1.Items.Add(CombName(i))
        Next

        SelectCombFrm.ShowDialog()

        If SelectCombFlag = False Then GoTo ExitSub
        If OpenFileErrorFlag = True Then GoTo exitsub

        Dim SelectedFrameCount As Integer
        '因中途有清除select 此行不執行
        'ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
        For i = 0 To NumberItems - 1
            If ObjectType(i) = 2 Then
                SelectedFrameCount += 1
            End If
        Next


        MsgBox(SelectedFrameCount.ToString + "  Frames Selected", , "Deflection Check")

        Dim UseVNodeFlag As Boolean = True
        If JointSequenceFile = "" Then

            UseVNodeFlag = False
            GoTo NoVNodeFile
        End If

        '========20130227=============================================
        FileOpen(25, JointSequenceFile, OpenMode.Input)
        Dim JointDataString As String
        Dim JointTableExistFlag As Boolean = False
        Dim AutoMeshPointCounter As Integer = 0

        Do Until EOF(25)
            JointDataString = LineInput(25)
            If JointDataString.Contains("Objects And Elements - Joints") Then JointTableExistFlag = True

            If JointDataString.Contains("~") Then
                AutoMeshPointCounter += 1
            End If
        Loop
        FileClose(25)

        FileOpen(26, JointSequenceFile, OpenMode.Input)

        Dim JointXYZ(AutoMeshPointCounter, 3) As String
        Dim JointDataTemp(4) As String

        Dim charseparators() As Char = " "
        Dim VNodeCounter As Integer = 0

        Do Until EOF(26)

            JointDataString = LineInput(26)
            If JointDataString.Contains("~") Then
                JointDataTemp = JointDataString.Split(charseparators, StringSplitOptions.RemoveEmptyEntries)

                'JointDataTemp(0) = JointDataTemp(0)
                JointXYZ(VNodeCounter, 0) = JointDataTemp(0)
                JointXYZ(VNodeCounter, 1) = JointDataTemp(1)
                JointXYZ(VNodeCounter, 2) = JointDataTemp(2)
                JointXYZ(VNodeCounter, 3) = JointDataTemp(3)
                VNodeCounter += 1

            End If

        Loop
        '====================================================================

        FileClose(26)


NoVNodeFile:

        FileClose(15)
        FileOpen(15, ReportFileName, OpenMode.Output)


        PrintLine(15, "")
        PrintLine(15, "")
        PrintLine(15, "      ======================================================")
        PrintLine(15, "                    Deflection Check Summary")
        PrintLine(15, "      ======================================================")
        PrintLine(15, "")

        If DefChkCriteria = True Then
            PrintLine(15, "Deflection Limit :   L/" + DeflectionCriteria.ToString)
        Else
            PrintLine(15, "Deflection Limit :  " + DeflectionCriteria.ToString + "   mm")
        End If


        PrintLine(15, "")


        '2015/07/15 format change for P1 Project===========

        Dim TitleLine As String = ""
        If CheckDeflectionRelative = True Then
            TitleLine = "   Frame No./    Section Name    /CTRL Comb/ Deflection (Relative) / OK-NG / Critical Node /Node List"

        Else
            TitleLine = "   Frame No./    Section Name    /CTRL Comb/ Deflection (Absulate) / OK-NG / Critical Node /Node List"
        End If
        'Dim TitleLine As String = "   Frame No./    Section Name    /CTRL Comb/ Deflection / ******** / OK-NG / Critical Node /Node List"

        If ShowSectName = False Then
            TitleLine = TitleLine.Replace("    Section Name    ", "********************")
        End If
        If ShowCriNode = False Then
            TitleLine = TitleLine.Replace(" Critical Node ", "***************")
        End If
        If ShowNodeList = False Then
            TitleLine = TitleLine.Replace("Node List", "*********")
        End If

        PrintLine(15, TitleLine)
        '==================================================

        'If ShowNodeList = True Then
        '    PrintLine(15, "   Frame No./CTRL Comb/ Deflection / ******** / OK-NG / Critical Node /Node List")
        'Else
        '    PrintLine(15, "   Frame No./CTRL Comb/ Deflection / ******** / OK-NG / Critical Node /")
        'End If
        'PrintLine(15, "                        mm          mm")
        PrintLine(15, "")
        PrintLine(15, " ====================================================================================================")

        Dim OKValue As Integer = DeflectionCriteria
        Dim OKFlag As String
        Dim NodeList As String
        Dim VoidNumCompare As String
        Dim Point1, Point2 As String

        Dim SkipCount As Integer

        For m = 0 To NumberItems - 1

            If ObjectType(m) <> 2 Then
                SkipCount += 1
                GoTo NextItem
            End If

            'get NumSegs
            ret = SapModel.FrameObj.GetAutoMesh(ObjectName(m), AutoMesh, AutoMeshAtPoints, AutoMeshAtLines, NumSegs, AutoMeshMaxLength)
            ret = SapModel.FrameObj.GetPoints(ObjectName(m), Point1, Point2)

            'VoidPointCounter = 1
            If NumSegs < 2 Or AutoMesh = False Then
                NumSegs = 0
                MsgBox("Frame  " + ObjectName(m) + " segement less than 2 please check!", MsgBoxStyle.Critical, "Deflection Check")

            End If

            NodeList = ""
            MaxNodeFlag = 0
            MaxCombFlag = 0

            MaxResult = 999999
            For j = 0 To CombList.Count - 1



                ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
                ret = SapModel.Results.Setup.SetComboSelectedForOutput(CombList(j))

                Elm = Nothing


                ret = SapModel.Results.JointDispl(ObjectName(m), eItemTypeElm.GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)

                ret = SapModel.Results.JointDisplAbs(ObjectName(m), eItemTypeElm.GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)


                '==========2014/01/03 避免Group name 沒變更的問題
                If Elm Is Nothing Then
                    '========20140715新增======================
                    Dim OBName As String = GroupDict(ObjectName(m)).Replace(",", "")
                    ret = SapModel.Results.JointDispl(OBName, eItemTypeElm.GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)

                    ret = SapModel.Results.JointDisplAbs(OBName, eItemTypeElm.GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)

                    '=====================================
                End If

                'ret = SapModel.Results.JointDispl("1382-2", eItemTypeElm.GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)


                If ret = 1 Then
                    MsgBox("Error Occurred when try to get Elm() Data" & vbCrLf & "Frame ID : " & vbCrLf & ObjectName(j))
                    GoTo ExitSub
                End If

                'ret = SapModel.PointObj.GetCoordCartesian(Obj(0), X1, Y1, Z1)

                ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)


                SPt.X = X1
                SPt.Y = Y1
                SPt.Z = Z1

                For i = 0 To Obj.Count - 1
                    If Obj(i) = Point1 Then
                        SDef = U3(i)
                        Exit For
                    End If
                Next


                'ret = SapModel.PointObj.GetCoordCartesian(Obj(1), X2, Y2, Z2)
                ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)

                EPt.X = X2
                EPt.Y = Y2
                EPt.Z = Z2
                For i = 0 To Obj.Count - 1
                    If Obj(i) = Point2 Then
                        EDef = U3(i)
                        Exit For
                    End If
                Next

                Dim VoidPointCounterByNumber As Integer = 0
                For i = 0 To Elm.Count - 1
                    If Elm(i).Contains("~") Then
                        VoidPointCounterByNumber += 1
                    End If

                Next



                'MaxResult = 999999
                VoidPointCounter = 1

                'For i = 2 To Elm.Count - 1
                For i = 0 To Elm.Count - 1

                    '=========filter column
                    If Math.Abs(Z1 - Z2) > 2500 Then Exit For
                    '======================

                    If Elm(i).Contains("~") And UseVNodeFlag = False Then
                        MPt.X = X1 + (X2 - X1) / (VoidPointCounterByNumber + 1) * VoidPointCounter
                        MPt.Y = Y1 + (Y2 - Y1) / (VoidPointCounterByNumber + 1) * VoidPointCounter
                        MPt.Z = Z1 + (Z2 - Z1) / (VoidPointCounterByNumber + 1) * VoidPointCounter
                        MDef = U3(i)
                        VoidPointCounter += 1
                        Result = fraction(SPt, EPt, MPt, SDef, EDef, MDef)
                    ElseIf Elm(i).Contains("~") And UseVNodeFlag = True Then
                        '=================================================
                        For V = 0 To VNodeCounter - 1
                            If JointXYZ(V, 0) = Elm(i) Then
                                MPt.X = JointXYZ(V, 1)
                                MPt.Y = JointXYZ(V, 2)
                                MPt.Z = JointXYZ(V, 3)
                                MDef = U3(i)
                                VoidPointCounter += 1
                                Result = fraction(SPt, EPt, MPt, SDef, EDef, MDef)
                                Exit For
                            End If
                        Next
                        '=================================================
                    Else

                        If Obj(i) <> Point1 And Obj(i) <> Point2 Then
                            ret = SapModel.PointObj.GetCoordCartesian(Obj(i), X3, Y3, Z3)
                            MPt.X = X3
                            MPt.Y = Y3
                            MPt.Z = Z3
                            MDef = U3(i)  'U3 2~n
                            Result = fraction(SPt, EPt, MPt, SDef, EDef, MDef)
                        Else
                            Result = 1000000
                        End If


                    End If

                    If 1 / MaxResult < 1 / Result Then
                        MaxResult = Result
                        MaxNodeFlag = i
                        MaxCombFlag = j
                    End If
                Next

            Next

            If DefChkCriteria = False Then MaxResult = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5 / MaxResult


            If DefChkCriteria = True Then
                If MaxResult > OKValue Then
                    OKFlag = "OK"
                Else
                    OKFlag = "NG"
                End If
            Else
                If MaxResult > OKValue Then
                    OKFlag = "NG"
                Else
                    OKFlag = "OK"
                End If
            End If



            For k = 0 To Elm.Count - 1
                NodeList = NodeList + " " + Elm(k)
            Next

            If ShowNodeList = False Then
                NodeList = ""
            End If

            If VoidPointCounter <> NumSegs And UseVNodeFlag = False Then
                '不顯示重複的node資料
                'VoidNumCompare = " (" + (NumSegs - VoidPointCounter).ToString + " Points Overlap" + ")"
                VoidNumCompare = ""
            Else
                VoidNumCompare = ""
            End If



            'If DefChkCriteria = True Then
            '    PrintLine(15, "    " + ObjectName(m), Microsoft.VisualBasic.TAB(13), CombList(MaxCombFlag), Microsoft.VisualBasic.TAB(25), "L/" + CStr(CInt(MaxResult)) + "   ", Microsoft.VisualBasic.TAB(48), OKFlag, Microsoft.VisualBasic.TAB(60), Elm(MaxNodeFlag), Microsoft.VisualBasic.TAB(68), "  |  " + NodeList + VoidNumCompare)
            'Else
            '    PrintLine(15, "    " + ObjectName(m), Microsoft.VisualBasic.TAB(13), CombList(MaxCombFlag), Microsoft.VisualBasic.TAB(25), " " + CStr(Format(CDbl(MaxResult), "0.00")) + "   ", Microsoft.VisualBasic.TAB(48), OKFlag, Microsoft.VisualBasic.TAB(60), Elm(MaxNodeFlag), Microsoft.VisualBasic.TAB(68), "  |  " + NodeList + VoidNumCompare)
            'End If


            '2015/07/15 format change for P1 Project===========
            Dim ResultString As String = Space(100)
            ResultString = ResultString.Insert(5, ObjectName(m))
            Dim SectionName As String
            Dim PropName As String
            Dim SAuto As String

            ret = SapModel.FrameObj.GetSection(ObjectName(m), SectionName, SAuto)



            If ShowSectName = True Then
                ResultString = ResultString.Insert(15, SectionName)
            Else
                ResultString = ResultString
            End If

            ResultString = ResultString.Insert(36, CombList(MaxCombFlag))

            If DefChkCriteria = True Then
                ResultString = ResultString.Insert(47, "L/" + CStr(CInt(MaxResult)))
            Else
                ResultString = ResultString.Insert(47, CStr(Format(CDbl(MaxResult), "0.00")))
            End If

            ResultString = ResultString.Insert(70, OKFlag)


            If ShowCriNode = True Then
                ResultString = ResultString.Insert(82, Elm(MaxNodeFlag))
            End If

            If ShowNodeList = True Then
                ResultString = ResultString.Insert(93, NodeList)
            End If

            PrintLine(15, ResultString.TrimEnd)

            '==================================================


NextItem: Next


        MsgBox("Complete" & vbCrLf & ReportFileName, , "Deflection Check")

ExitSub:
        FileClose(15)
        ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Deflection Check")
    End Sub

    '建立Group 與 Frame之對應關係
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


    '撓度檢查用Function
    Public Function fraction(ByRef StartPoint As Point3D, ByRef EndPoint As Point3D, ByVal MidPoint As Point3D, ByVal SDef As Double, ByVal EDef As Double, ByVal MDef As Double) As Double

        Dim length1, length2 As Double
        Dim linearDef As Double
        Dim RelDef As Double
        Dim AbsDef As Double

        If CheckDeflectionRelative = True Then
            length1 = ((StartPoint.X - MidPoint.X) ^ 2 + (StartPoint.Y - MidPoint.Y) ^ 2 + (StartPoint.Z - MidPoint.Z) ^ 2) ^ 0.5
            length2 = ((EndPoint.X - MidPoint.X) ^ 2 + (EndPoint.Y - MidPoint.Y) ^ 2 + (EndPoint.Z - MidPoint.Z) ^ 2) ^ 0.5

            linearDef = EDef + (SDef - EDef) / (length1 + length2) * length2
            RelDef = MDef - linearDef
            fraction = Math.Abs(1 / (RelDef / (length1 + length2)))
        End If

        If CheckDeflectionAbsolute = True Then
            length1 = ((StartPoint.X - MidPoint.X) ^ 2 + (StartPoint.Y - MidPoint.Y) ^ 2 + (StartPoint.Z - MidPoint.Z) ^ 2) ^ 0.5
            length2 = ((EndPoint.X - MidPoint.X) ^ 2 + (EndPoint.Y - MidPoint.Y) ^ 2 + (EndPoint.Z - MidPoint.Z) ^ 2) ^ 0.5

            'linearDef = EDef + (SDef - EDef) / (length1 + length2) * length2
            'RelDef = MDef - linearDef
            AbsDef = MDef
            fraction = Math.Abs(1 / (AbsDef / (length1 + length2)))
        End If



    End Function

    '撓度檢查用Data Structure
    Structure Point3D
        Dim X As Double
        Dim Y As Double
        Dim Z As Double
    End Structure

    '水池自動化
    Private Sub Create_Pool(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim ret As Long
        Dim x() As Double
        Dim y() As Double
        Dim z() As Double
        Dim Name As String

        ret = SapModel.InitializeNewModel(eUnits.Ton_m_C)
        ret = SapModel.File.NewBlank

        '==========set coordinate system=======

        SapModel.SetPresentUnits(eUnits.Ton_m_C)





        Dim inputfrm As New PoolDataForm
        inputfrm.ShowDialog()

        If inputfrm.DialogResult <> DialogResult.OK Then GoTo endSub


        Dim gradualWallName As String
        Dim gradualWallThk As Double

        For i = 1 To Divide_Z
            gradualWallName = "Wall_" & i
            gradualWallThk = TopWallThk + (i - 1) * (BottomWallThk - TopWallThk) / (Divide_Z - 1)

            gradualWallThk = Math.Round(gradualWallThk, 5)

            ret = SapModel.PropArea.SetShell(gradualWallName, 2, "4000Psi", 0, gradualWallThk, gradualWallThk, , "Created by Auto Pool")

        Next


        ret = SapModel.PropArea.SetShell("Bottom", 2, "4000Psi", 0, BottomSlabThk, BottomSlabThk, , "Created by Auto Pool")

        Dim Value() As Double
        ReDim Value(9)
        For i = 0 To 9
            Value(i) = 1
        Next i
        Value(8) = 0
        ret = SapModel.PropArea.SetModifiers("Bottom", Value)




        If WithTopSlab = True Then
            ret = SapModel.PropArea.SetShell("Top", 2, "4000Psi", 0, TopSlabThk, TopSlabThk, , "Created by Auto Pool")
        End If



        ReDim x(3)
        ReDim y(3)
        ReDim z(3)
        '======Create Bottom==========
        x(0) = 0 : y(0) = 0
        x(1) = Pool_Length : y(1) = 0
        x(2) = Pool_Length : y(2) = Pool_Width
        x(3) = 0 : y(3) = Pool_Width

        ret = SapModel.AreaObj.AddByCoord(4, x, y, z, Name, "Bottom")
        'refresh view
        ret = SapModel.View.RefreshView(0, True)

        Dim AreaName() As String

        'divide area object
        ret = SapModel.EditArea.Divide(Name, 1, Divide_X * Divide_Y, AreaName, Divide_X, Divide_Y)
        '=========Bottom Spring=========
        ret = SapModel.SelectObj.PlaneXY("1")
        Dim Vec() As Double
        ret = SapModel.AreaObj.SetSpring(Name, 1, VerSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)



        '=========







        Dim WallType As String
        '======Create 1st Wall==========
        For i = 1 To Divide_Z

            WallType = "Wall_" + CStr(Divide_Z + 1 - i)



            x(0) = 0 : y(0) = Pool_Width : z(0) = Pool_Height / Divide_Z * (i - 1)
            x(1) = Pool_Length : y(1) = Pool_Width : z(1) = Pool_Height / Divide_Z * (i - 1)
            x(2) = Pool_Length : y(2) = Pool_Width : z(2) = Pool_Height / Divide_Z * i
            x(3) = 0 : y(3) = Pool_Width : z(3) = Pool_Height / Divide_Z * i

            ret = SapModel.AreaObj.AddByCoord(4, x, y, z, Name, WallType)
            ret = SapModel.EditArea.Divide(Name, 1, Divide_X, AreaName, Divide_X, 1)
        Next

        'ret = SapModel.EditArea.Divide(Name, 1, Divide_X * Divide_Z, AreaName, Divide_X, Divide_Z)
        '=========Create joint Pattern
        'ret = SapModel.PatternDef.SetPattern("Impulsive")
        'ret = SapModel.SelectObj.PlaneXZ("4")
        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive", 0, 0, -1 * (WaterPressure_Bottom - WaterPressure_Top) / Pool_Height, WaterPressure_Bottom, eItemType.SelectedObjects, 0, False)
        'ret = SapModel.LoadPatterns.Add("Impulsive", eLoadPatternType.LTYPE_WATERLOADPRESSURE)
        'ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Impulsive", -2, 1, "Impulsive", False, eItemType.SelectedObjects)

        'ret = SapModel.PatternDef.SetPattern("Convective")
        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Convective", 0, 0, -1 * (WaterPressure2_Bottom - WaterPressure2_Top) / Pool_Height, WaterPressure2_Bottom, eItemType.SelectedObjects, 0, False)
        'ret = SapModel.LoadPatterns.Add("Convective", eLoadPatternType.LTYPE_WATERLOADPRESSURE)
        'ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Convective", -2, 1, "Convective", False, eItemType.SelectedObjects)



        'ret = SapModel.AreaObj.SetSpring(Name, 1, HoriSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)





        '======Create 2nd Wall==========
        For i = 1 To Divide_Z

            WallType = "Wall_" + CStr(Divide_Z + 1 - i)

            x(0) = Pool_Length : y(0) = 0 : z(0) = Pool_Height / Divide_Z * (i - 1)
            x(1) = 0 : y(1) = 0 : z(1) = Pool_Height / Divide_Z * (i - 1)
            x(2) = 0 : y(2) = 0 : z(2) = Pool_Height / Divide_Z * i
            x(3) = Pool_Length : y(3) = 0 : z(3) = Pool_Height / Divide_Z * i

            ret = SapModel.AreaObj.AddByCoord(4, x, y, z, Name, WallType)
            ret = SapModel.EditArea.Divide(Name, 1, Divide_X, AreaName, Divide_X, 1)

        Next

        'ret = SapModel.SelectObj.PlaneXZ("1")
        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive", 0, 0, -1 * (WaterPressure_Bottom - WaterPressure_Top) / Pool_Height, WaterPressure_Bottom, eItemType.SelectedObjects, 0, False)
        ''ret = SapModel.LoadPatterns.Add("Impulsive", eLoadPatternType.LTYPE_WAVE)
        'ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Impulsive", -2, 1, "Impulsive", False, eItemType.SelectedObjects)

        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Convective", 0, 0, -1 * (WaterPressure2_Bottom - WaterPressure2_Top) / Pool_Height, WaterPressure2_Bottom, eItemType.SelectedObjects, 0, False)
        'ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Convective", -2, 1, "Convective", False, eItemType.SelectedObjects)

        'ret = SapModel.AreaObj.SetSpring(Name, 1, HoriSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)



        '' ''======Create 3rd Wall==========

        For i = 1 To Divide_Z

            WallType = "Wall_" + CStr(Divide_Z + 1 - i)

            x(0) = 0 : y(0) = 0 : z(0) = Pool_Height / Divide_Z * (i - 1)
            x(1) = 0 : y(1) = Pool_Width : z(1) = Pool_Height / Divide_Z * (i - 1)
            x(2) = 0 : y(2) = Pool_Width : z(2) = Pool_Height / Divide_Z * i
            x(3) = 0 : y(3) = 0 : z(3) = Pool_Height / Divide_Z * i

            ret = SapModel.AreaObj.AddByCoord(4, x, y, z, Name, WallType)
            ret = SapModel.EditArea.Divide(Name, 1, Divide_Y, AreaName, Divide_Y, 1)


        Next


        'ret = SapModel.SelectObj.PlaneYZ("1")

        'WaterPressure_Bottom = 20
        'WaterPressure_Top = 12

        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive", 0, 0, -1 * (WaterPressure_Bottom - WaterPressure_Top) / Pool_Height, WaterPressure_Bottom, eItemType.SelectedObjects, 0, True)
        'ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Impulsive", -2, 1, "Impulsive", False, eItemType.SelectedObjects)

        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Convective", 0, 0, -1 * (WaterPressure2_Bottom - WaterPressure2_Top) / Pool_Height, WaterPressure2_Bottom, eItemType.SelectedObjects, 0, False)
        'ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Convective", -2, 1, "Convective", False, eItemType.SelectedObjects)

        'ret = SapModel.AreaObj.SetSpring(Name, 1, HoriSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)



        '======Create 4th Wall==========

        For i = 1 To Divide_Z

            WallType = "Wall_" + CStr(Divide_Z + 1 - i)

            x(0) = Pool_Length : y(0) = Pool_Width : z(0) = Pool_Height / Divide_Z * (i - 1)
            x(1) = Pool_Length : y(1) = 0 : z(1) = Pool_Height / Divide_Z * (i - 1)
            x(2) = Pool_Length : y(2) = 0 : z(2) = Pool_Height / Divide_Z * i
            x(3) = Pool_Length : y(3) = Pool_Width : z(3) = Pool_Height / Divide_Z * i

            ret = SapModel.AreaObj.AddByCoord(4, x, y, z, Name, WallType)
            ret = SapModel.EditArea.Divide(Name, 1, Divide_Y, AreaName, Divide_Y, 1)


        Next



        'ret = SapModel.SelectObj.PlaneYZ("2")
        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive", 0, 0, -1 * (WaterPressure_Bottom - WaterPressure_Top) / Pool_Height, WaterPressure_Bottom, eItemType.SelectedObjects, 0, False)
        'ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Impulsive", -2, 1, "Impulsive", False, eItemType.SelectedObjects)

        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Convective", 0, 0, -1 * (WaterPressure2_Bottom - WaterPressure2_Top) / Pool_Height, WaterPressure2_Bottom, eItemType.SelectedObjects, 0, False)
        'ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Convective", -2, 1, "Convective", False, eItemType.SelectedObjects)

        'ret = SapModel.AreaObj.SetSpring(Name, 1, HoriSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)



        '' ''======Create Top Slab=========
        If WithTopSlab = True Then

            x(0) = 0 : y(0) = 0 : z(0) = Pool_Height
            x(1) = Pool_Length : y(1) = 0 : z(1) = Pool_Height
            x(2) = Pool_Length : y(2) = Pool_Width : z(2) = Pool_Height
            x(3) = 0 : y(3) = Pool_Width : z(3) = Pool_Height


            ret = SapModel.AreaObj.AddByCoord(4, x, y, z, Name, "Top")
            ret = SapModel.EditArea.Divide(Name, 1, Divide_X * Divide_Y, AreaName, Divide_X, Divide_Y)
        End If


        '===========Apply force to wall==========
        '===========1st wall
        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.PlaneXZ("4")

        ret = SapModel.SelectObj.CoordinateRange(0, Pool_Length, 0, Pool_Width, 0.01 + WaterDepth, Pool_Height, True, , True)

        ret = SapModel.PatternDef.SetPattern("WaterPressure")
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "WaterPressure", 0, 0, -1, WaterDepth, eItemType.SelectedObjects, 0, True)
        ret = SapModel.LoadPatterns.Add("WaterPressure", eLoadPatternType.WaterloadPressure)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "WaterPressure", -2, 1, "WaterPressure", True, eItemType.SelectedObjects)

        ret = SapModel.PatternDef.SetPattern("Impulsive_Y")
        'ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive", 0, 0, -1 * (WaterPressure_Bottom - WaterPressure_Top) / Pool_Height, WaterPressure_Bottom, eItemType.SelectedObjects, 0, False)
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive_Y", 0, 0, -1 * (WaterPressure_Bottom - WaterPressure_Top) / WaterDepth, WaterPressure_Bottom, eItemType.SelectedObjects, 0, False)

        ret = SapModel.LoadPatterns.Add("Impulsive_Y", eLoadPatternType.WaterloadPressure)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Impulsive_Y", -2, 1, "Impulsive_Y", False, eItemType.SelectedObjects)
        ret = SapModel.PatternDef.SetPattern("Convective_Y")
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "Convective_Y", 0, 0, -1 * (WaterPressure2_Bottom - WaterPressure2_Top) / WaterDepth, WaterPressure2_Bottom, eItemType.SelectedObjects, 0, False)
        ret = SapModel.LoadPatterns.Add("Convective_Y", eLoadPatternType.WaterloadPressure)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Convective_Y", -2, 1, "Convective_Y", False, eItemType.SelectedObjects)

        'ret = SapModel.SelectObj.ClearSelection
        'ret = SapModel.SelectObj.PlaneXZ("4")
        ret = SapModel.AreaObj.SetSpring(Name, 1, HoriSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)
        'ret = SapModel.SelectObj.ClearSelection
        'ret = SapModel.SelectObj.CoordinateRange(0, Pool_Length, 0, Pool_Width, 0.01 + WaterDepth, Pool_Height, False, , True)

        'ret = SapModel.AreaObj.DeleteSpring(Name, eItemType.SelectedObjects)




        '===========2nd wall
        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.PlaneXZ("1")

        ret = SapModel.SelectObj.CoordinateRange(0, Pool_Length, 0, Pool_Width, 0.01 + WaterDepth, Pool_Height, True, , True)

        'ret = SapModel.PatternDef.SetPattern("WaterPressure")
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "WaterPressure", 0, 0, -1, WaterDepth, eItemType.SelectedObjects, 0, True)
        'ret = SapModel.LoadPatterns.Add("WaterPressure", eLoadPatternType.LTYPE_WATERLOADPRESSURE)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "WaterPressure", -2, 1, "WaterPressure", True, eItemType.SelectedObjects)


        ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive_Y", 0, 0, -1 * (WaterPressure_Bottom - WaterPressure_Top) / WaterDepth, WaterPressure_Bottom, eItemType.SelectedObjects, 0, False)
        'ret = SapModel.LoadPatterns.Add("Impulsive", eLoadPatternType.LTYPE_WAVE)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Impulsive_Y", -2, -1, "Impulsive_Y", False, eItemType.SelectedObjects)
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "Convective_Y", 0, 0, -1 * (WaterPressure2_Bottom - WaterPressure2_Top) / WaterDepth, WaterPressure2_Bottom, eItemType.SelectedObjects, 0, False)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Convective_Y", -2, -1, "Convective_Y", False, eItemType.SelectedObjects)
        ret = SapModel.AreaObj.SetSpring(Name, 1, HoriSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)


        '===========3rd wall
        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.PlaneYZ("1")

        ret = SapModel.SelectObj.CoordinateRange(0, Pool_Length, 0, Pool_Width, 0.01 + WaterDepth, Pool_Height, True, , True)

        'ret = SapModel.PatternDef.SetPattern("WaterPressure")
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "WaterPressure", 0, 0, -1, WaterDepth, eItemType.SelectedObjects, 0, True)
        'ret = SapModel.LoadPatterns.Add("WaterPressure", eLoadPatternType.LTYPE_WATERLOADPRESSURE)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "WaterPressure", -2, 1, "WaterPressure", True, eItemType.SelectedObjects)


        'WaterPressure_Bottom = 20
        'WaterPressure_Top = 12
        ret = SapModel.PatternDef.SetPattern("Impulsive_X")
        ret = SapModel.PatternDef.SetPattern("Convective_X")
        ret = SapModel.LoadPatterns.Add("Impulsive_X", eLoadPatternType.WaterloadPressure)
        ret = SapModel.LoadPatterns.Add("Convective_X", eLoadPatternType.WaterloadPressure)

        ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive_X", 0, 0, -1 * (WaterPressure3_Bottom - WaterPressure3_Top) / WaterDepth, WaterPressure3_Bottom, eItemType.SelectedObjects, 0, False)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Impulsive_X", -2, -1, "Impulsive_X", False, eItemType.SelectedObjects)
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "Convective_X", 0, 0, -1 * (WaterPressure4_Bottom - WaterPressure4_Top) / WaterDepth, WaterPressure4_Bottom, eItemType.SelectedObjects, 0, False)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Convective_X", -2, -1, "Convective_X", False, eItemType.SelectedObjects)
        ret = SapModel.AreaObj.SetSpring(Name, 1, HoriSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)


        '===========4th wall
        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.PlaneYZ("2")

        ret = SapModel.SelectObj.CoordinateRange(0, Pool_Length, 0, Pool_Width, 0.01 + WaterDepth, Pool_Height, True, , True)

        'ret = SapModel.PatternDef.SetPattern("WaterPressure")
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "WaterPressure", 0, 0, -1, WaterDepth, eItemType.SelectedObjects, 0, True)
        'ret = SapModel.LoadPatterns.Add("WaterPressure", eLoadPatternType.LTYPE_WATERLOADPRESSURE)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "WaterPressure", -2, 1, "WaterPressure", True, eItemType.SelectedObjects)



        ret = SapModel.PointObj.SetPatternByXYZ(Name, "Impulsive_X", 0, 0, -1 * (WaterPressure3_Bottom - WaterPressure3_Top) / WaterDepth, WaterPressure3_Bottom, eItemType.SelectedObjects, 0, False)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Impulsive_X", -2, 1, "Impulsive_X", False, eItemType.SelectedObjects)
        ret = SapModel.PointObj.SetPatternByXYZ(Name, "Convective_X", 0, 0, -1 * (WaterPressure4_Bottom - WaterPressure4_Top) / WaterDepth, WaterPressure4_Bottom, eItemType.SelectedObjects, 0, False)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "Convective_X", -2, 1, "Convective_X", False, eItemType.SelectedObjects)
        ret = SapModel.AreaObj.SetSpring(Name, 1, HoriSoil_K, 2, "", -1, 1, 3, False, Vec, 0, False, "Local", eItemType.SelectedObjects)


        '==========Bottom Slab Water Pressure
        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.PlaneXY("1")

        'ret = SapModel.PatternDef.SetPattern("WaterPressure")
        'ret = SapModel.LoadPatterns.Add("WaterPressure", eLoadPatternType.LTYPE_WATERLOADPRESSURE)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "WaterPressure", -2, WaterDepth, , True, eItemType.SelectedObjects)

        ret = SapModel.PatternDef.SetPattern("VerSeismicLoad")
        ret = SapModel.LoadPatterns.Add("VerSeismicLoad", eLoadPatternType.Quake)
        ret = SapModel.AreaObj.SetLoadSurfacePressure(Name, "VerSeismicLoad", -2, WaterDepth * VerSeismicCoeff, , True, eItemType.SelectedObjects)


        '===========Apply force end
        ret = SapModel.SelectObj.ClearSelection

        MsgBox("Pool Complete", MsgBoxStyle.Information, "Auto Create Pool")

endSub: ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Create Pool")
    End Sub

    'Plug-in版本檢查
    Private Sub Ver_Check(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        MsgBox("This is SAP 2000 V20.X Plug-in  ", MsgBoxStyle.Information)  '每次需修改

        Dim UsedVer As String
        UsedVer = "20.1"   '每次需修改

        FileOpen(100, "\\c1199\util\SAP2000_Plugin\V20\VersionCheck.txt", OpenMode.Input)
        'Dim fileReader As String

        'fileReader = My.Computer.FileSystem.ReadAllText("N:\SAP2000_Plugin\V16.0.2\VersionCheck.txt")

        'Dim Sr As IO.StreamReader

        'fileReader = Sr.ReadToEnd(


        Dim CurrentVersion As String
        CurrentVersion = LineInput(100)

        If UsedVer <> CurrentVersion Then
            MsgBox("Your Version :  " & UsedVer & vbCrLf & "Current Version : " & CurrentVersion & vbCrLf & vbCrLf & _
                    "Please Update SAP2000 Plug-in", MsgBoxStyle.Exclamation)
        Else
            MsgBox("Plug-in Version Check   :   OK! ")
        End If

        FileClose(100)

        ISapPlugin.Finish(0)
    End Sub

    '功能說明
    Private Sub Open_Manual(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Dim UserManual As New Manual

        UserManual.Show()


    End Sub

    '柱自動分段
    Private Sub Split_Column(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        Dim ret As Long
        Dim ParallelTo(5) As Boolean
        Dim Num As Long
        Dim NewName() As String
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim memberCounter, NewMember As Integer

        'ret = SapModel.SelectObj.All
        ParallelTo(2) = True
        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.LinesParallelToCoordAxis(ParallelTo)
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        If NumberItems = 0 Then
            GoTo Fin
        End If

        For i = 0 To ObjectType.Count - 1
            If ObjectType(i) = 2 Then    '2 = Frame object
                ret = SapModel.SelectObj.ClearSelection

                ret = SapModel.PointObj.SetSelected("All", True, eItemType.Group)
                ret = SapModel.FrameObj.SetSelected(ObjectName(i), True)
                ret = SapModel.EditFrame.DivideAtIntersections(ObjectName(i), Num, NewName)
                memberCounter += 1
            End If
        Next

        ret = SapModel.SelectObj.ClearSelection
        ret = SapModel.SelectObj.LinesParallelToCoordAxis(ParallelTo)
        ret = SapModel.SelectObj.GetSelected(NewMember, ObjectType, ObjectName)

Fin:
        MsgBox(memberCounter & "  Members Selected" & _
               vbCrLf & NewMember - NumberItems & "  New Members Created", MsgBoxStyle.Information, "Complete")

        ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Split Column")
    End Sub

    'L/30 梁深檢查
    Private Sub Check_Beam_Depth(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        Dim SDBFilePath As String
        Dim resultFile As String

        SDBFilePath = My.Computer.FileSystem.CurrentDirectory
        resultFile = SDBFilePath + "\Beam Depth Check Result.txt"


        Dim ret As Long
        Dim ParallelTo(5) As Boolean
        Dim Num As Long
        Dim NewName() As String
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim memberCounter, NewMember As Integer

        Dim PropName As String
        Dim SAuto As String


        'ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
        ret = SapModel.SelectObj.ClearSelection

        ParallelTo(0) = True
        ParallelTo(1) = True
        ret = SapModel.SelectObj.LinesParallelToCoordAxis(ParallelTo)
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        Dim FileName As String
        Dim MatProp As String
        Dim t3 As Double
        Dim t2 As Double
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
        Dim Color As Long
        Dim Notes As String
        Dim GUID As String

        Dim tf As Double
        Dim tw As Double
        Dim t2b As Double
        Dim tfb As Double

        Dim Point1 As String
        Dim Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double
        Dim memberLength As Double
        Dim UnknowFrame(NumberItems - 1) As String
        Dim ProblemFrame(NumberItems - 1) As String

        For i = 0 To NumberItems - 1
            ret = SapModel.FrameObj.GetSection(ObjectName(i), PropName, SAuto)
            t3 = 0
            If PropName.Contains("H") Or PropName.Contains("FSEC") Then
                ret = SapModel.PropFrame.GetISection(PropName, FileName, MatProp, t3, t2, tf, tw, t2b, tfb, Color, Notes, GUID)
            ElseIf PropName.Contains("C") And Not PropName.Contains("RC") Then
                ret = SapModel.PropFrame.GetChannel(PropName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
            ElseIf PropName.Contains("T") Then
                ret = SapModel.PropFrame.GetTee(PropName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
            End If

            If t3 = 0 Then
                UnknowFrame(i) = ObjectName(i)
                If PropName.Contains("RC") Then
                    UnknowFrame(i) = UnknowFrame(i) + "(RC)"
                End If

            Else
                ret = SapModel.FrameObj.GetPoints(ObjectName(i), Point1, Point2)
                ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
                ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)
                memberLength = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5

                If t3 < memberLength / 30 Then
                    ProblemFrame(i) = ObjectName(i)

                End If
            End If

        Next

        Dim CantRecognizrList As String
        Dim UnknowCounter As Integer
        For i = 0 To NumberItems - 1
            If UnknowFrame(i) <> Nothing Then
                If Not UnknowFrame(i).Contains("RC") Then
                    CantRecognizrList = CantRecognizrList + UnknowFrame(i) + ","
                    UnknowCounter += 1
                End If
            End If
        Next

        If UnknowCounter > 0 Then
            MsgBox("Can't recognize these members :" & vbCrLf & CantRecognizrList & vbCrLf & vbCrLf _
                   & "Check Frame Definition", MsgBoxStyle.Question, "Check Beam Depth")
            ISapPlugin.Finish(0)
            Exit Sub
        End If


        ret = SapModel.SelectObj.ClearSelection
        Dim Result As String
        Dim ProblemFrameCounter As Integer
        For j = 0 To NumberItems - 1
            If ProblemFrame(j) <> Nothing Then
                Result = Result + ObjectName(j) + ","
                ret = SapModel.FrameObj.SetSelected(ObjectName(j), True)
                ProblemFrameCounter += 1
            End If
        Next

        Dim ProjectInfo As Long
        Dim Item() As String
        Dim Data() As String
        Dim ModelName As String = ""

        ret = SapModel.GetProjectInfo(ProjectInfo, Item, Data)
        If ProjectInfo = 0 Then GoTo InputName
        For i = 0 To Item.Count - 1
            If Item(i) = "Model Name" Then
                ModelName = Data(i)
            End If
        Next
InputName:

        ModelName = SapModel.GetModelFilename(False)

        If ModelName = "" Then
            ModelName = InputBox("Please Input Model Name : ", "Input")
            ret = SapModel.SetProjectInfo("Model Name", ModelName)
        End If


        FileOpen(60, resultFile, OpenMode.Output)

        PrintLine(60, "")
        PrintLine(60, "")
        PrintLine(60, "===============================================================")
        PrintLine(60, "                    Beam Depth Check Result                    ")
        PrintLine(60, "===============================================================")
        PrintLine(60, "")
        PrintLine(60, "")
        PrintLine(60, "Time       :  " & Now)
        PrintLine(60, "Model Name :  " & ModelName)

        If ProblemFrameCounter = 0 Then
            MsgBox("All Steel Members are OK!", , "Check Beam Depth")
            PrintLine(60, "")
            PrintLine(60, "All Steel Members Depth Are Greater Then Length/30 ")
            PrintLine(60, "")
            PrintLine(60, "                 ")
            PrintLine(60, "=      O K      =")
            PrintLine(60, "                 ")

        Else
            MsgBox("Steel Frame depth less then length/30  :" & vbCrLf & Result & _
                   vbCrLf & vbCrLf & ProblemFrameCounter & "   Frames selected", MsgBoxStyle.Critical, "Check Beam Depth")
            PrintLine(60, "")
            If ProblemFrameCounter = 1 Then
                PrintLine(60, ProblemFrameCounter & " Frame Depth Is Less Then Length/ 30 ")
            ElseIf ProblemFrameCounter > 2 Then
                PrintLine(60, ProblemFrameCounter & " Frames Depth Are Less Then Length/ 30 ")
            End If

            PrintLine(60, "")
            PrintLine(60, "Frame ID : ")
            PrintLine(60, Result)
            PrintLine(60, "")
            PrintLine(60, "                 ")
            PrintLine(60, "=    Caution    =")
            PrintLine(60, "                 ")
        End If

        MsgBox(resultFile, , "Check Result File")

        FileClose(60)
        Process.Start(resultFile)

        ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Check beam depth")
    End Sub

    '反力檔輸出
    Private Sub Output_Reaction(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        SapModel.SetPresentUnits(eUnits.Ton_m_C)
        Dim ret As Long
        'Dim PointItems As String
        Dim NumberResults As Integer
        Dim Obj() As String
        Dim Elm() As String
        Dim LoadCase() As String
        Dim StepType() As String
        Dim StepNum() As Double
        Dim F1() As Double
        Dim F2() As Double
        Dim F3() As Double
        Dim M1() As Double
        Dim M2() As Double
        Dim M3() As Double

        Dim NumberCombs As Integer
        'Dim Selected As Boolean
        'Dim NumberItems As Integer
        Dim CombName() As String
        Dim NumberItemsA As Integer
        Dim CaseNameA() As String
        Dim Status() As Integer
        Dim AnalyzedFlag As Boolean = False

        ret = SapModel.Analyze.GetCaseStatus(NumberItemsA, CaseNameA, Status)

        For i = 0 To NumberItemsA - 1
            If Status(i) = 4 Then
                AnalyzedFlag = True
            End If
        Next

        If AnalyzedFlag <> True Then
            MsgBox("Need To Run Analysis First                 " & vbCrLf & "End Program", , "Output Support Reaction")
            GoTo ExitSub
        End If

        'Dim NumberCases As Integer
        'Dim CaseName() As String

        'ret = SapModel.RespCombo.GetNameList(NumberCombs, CombName)

        'ret = SapModel.LoadCases.GetNameList(NumberCases, CaseName)
        'run analysis
        'ret = SapModel.File.Save("C:\SapAPI\x.sdb")
        'ret = SapModel.Analyze.RunAnalysis


        'Dim CaseName() As String
        'Dim Status() As Long
        'ret = SapModel.Analyze.GetCaseStatus(NumberItems, CaseName, Status)
        ret = SapModel.SelectObj.ClearSelection

        Dim DOF() As Boolean
        'select supported points
        ReDim DOF(5)
        DOF(0) = True
        DOF(1) = True
        DOF(2) = True
        DOF(3) = True
        DOF(4) = True
        DOF(5) = True
        ret = SapModel.SelectObj.SupportedPoints(DOF)

        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String


        'get selected objects
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        Dim SaveFileDia As New SaveFileDialog
        Dim ReaFileName As String
        Dim LraFileName As String

        SaveFileDia.FileName = "Reaction"
        SaveFileDia.DefaultExt = "rea"
        SaveFileDia.Filter = "Reaction file(*.rea) |*.rea"


        If SaveFileDia.ShowDialog() = DialogResult.OK Then
            ReaFileName = SaveFileDia.FileName
            LraFileName = ReaFileName.Replace(".rea", ".lra")
            FileOpen(10, SaveFileDia.FileName, OpenMode.Output)
            FileOpen(20, LraFileName, OpenMode.Output)
        Else
            GoTo exitsub
        End If



        'FileOpen(10, "V:\54833\Reaction.txt", OpenMode.Output)
        ret = SapModel.RespCombo.GetNameList(NumberCombs, CombName)
        Dim NumberCases As Integer
        Dim CaseName() As String

        ret = SapModel.LoadCases.GetNameList(NumberCases, CaseName)

        PrintLine(10, AlignS(NumberCombs, 0, 5), Microsoft.VisualBasic.TAB(9), "UNIT : Ton,m       FX        FY        FZ        MX        MY        MZ")
        PrintLine(20, AlignS(NumberCases, 0, 5), Microsoft.VisualBasic.TAB(9), "UNIT : Ton,m       FX        FY        FZ        MX        MY        MZ")

        Dim FileIndex As Integer
        Dim CombNameA() As String
        Dim CaseCount As Integer = 0

        For K = 0 To 1
            If K = 0 Then
                FileIndex = 10
            ElseIf K = 1 Then
                FileIndex = 20
                ret = SapModel.LoadCases.GetNameList(NumberCombs, CombNameA)
                Array.Clear(CombName, 0, CombName.Count)

                If CombNameA.Contains("MODAL") Then
                    ReDim CombName(NumberCombs - 2)
                Else
                    ReDim CombName(NumberCombs - 1)
                End If

                For i = 0 To NumberCombs - 1

                    If CombNameA(i) <> "MODAL" Then
                        CombName(CaseCount) = CombNameA(i)
                        CaseCount += 1
                    Else
                        NumberCombs = NumberCombs - 1
                    End If

                Next
            End If


            Dim FxMaxValue As Double = -10000
            Dim FxMinValue As Double = 10000
            Dim FxMaxnode As Integer
            Dim FxMinNode As Integer
            Dim FxMaxComb As String
            Dim FxMinComb As String
            Dim FyMaxValue As Double = -10000
            Dim FyMinValue As Double = 10000
            Dim FyMaxnode As Integer
            Dim FyMinNode As Integer
            Dim FyMaxComb As String
            Dim FyMinComb As String
            Dim FzMaxValue As Double = -10000
            Dim FzMinValue As Double = 10000
            Dim FzMaxnode As Integer
            Dim FzMinNode As Integer
            Dim FzMaxComb As String
            Dim FzMinComb As String
            Dim MxMaxValue As Double = -10000
            Dim MxMinValue As Double = 10000
            Dim MxMaxnode As Integer
            Dim MxMinNode As Integer
            Dim MxMaxComb As String
            Dim MxMinComb As String
            Dim MyMaxValue As Double = -10000
            Dim MyMinValue As Double = 10000
            Dim MyMaxnode As Integer
            Dim MyMinNode As Integer
            Dim MyMaxComb As String
            Dim MyMinComb As String
            Dim MzMaxValue As Double = -10000
            Dim MzMinValue As Double = 10000
            Dim MzMaxnode As Integer
            Dim MzMinNode As Integer
            Dim MzMaxComb As String
            Dim MzMinComb As String


            For j = 0 To NumberItems - 1
                For i = 0 To NumberCombs - 1
                    'clear all case and combo output selections
                    ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
                    'set case and combo output selections
                    If K = 0 Then
                        ret = SapModel.Results.Setup.SetComboSelectedForOutput(CombName(i))
                        ret = SapModel.Results.JointReact(ObjectName(j), eItemTypeElm.Element, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)

                    Else
                        ret = SapModel.Results.Setup.SetCaseSelectedForOutput(CombName(i))
                        If CombName(i) = "MODAL" Then
                            GoTo pass
                        End If
                        ret = SapModel.Results.JointReact(ObjectName(j), eItemTypeElm.Element, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)
                    End If


                    If FxMaxValue < F1(0) Then
                        FxMaxValue = F1(0)
                        FxMaxnode = ObjectName(j)
                        FxMaxComb = CombName(i)
                    End If

                    If FxMinValue > F1(0) Then
                        FxMinValue = F1(0)
                        FxMinNode = ObjectName(j)
                        FxMinComb = CombName(i)
                    End If


                    If FyMaxValue < F2(0) Then
                        FyMaxValue = F2(0)
                        FyMaxnode = ObjectName(j)
                        FyMaxComb = CombName(i)
                    End If

                    If FyMinValue > F2(0) Then
                        FyMinValue = F2(0)
                        FyMinNode = ObjectName(j)
                        FyMinComb = CombName(i)
                    End If

                    If FzMaxValue < F3(0) Then
                        FzMaxValue = F3(0)
                        FzMaxnode = ObjectName(j)
                        FzMaxComb = CombName(i)
                    End If

                    If FzMinValue > F3(0) Then
                        FzMinValue = F3(0)
                        FzMinNode = ObjectName(j)
                        FzMinComb = CombName(i)
                    End If

                    If MxMaxValue < M1(0) Then
                        MxMaxValue = M1(0)
                        MxMaxnode = ObjectName(j)
                        MxMaxComb = CombName(i)
                    End If

                    If MxMinValue > M1(0) Then
                        MxMinValue = M1(0)
                        MxMinNode = ObjectName(j)
                        MxMinComb = CombName(i)
                    End If

                    If MyMaxValue < M2(0) Then
                        MyMaxValue = M2(0)
                        MyMaxnode = ObjectName(j)
                        MyMaxComb = CombName(i)
                    End If

                    If MyMinValue > M2(0) Then
                        MyMinValue = M2(0)
                        MyMinNode = ObjectName(j)
                        MyMinComb = CombName(i)
                    End If

                    If MzMaxValue < M3(0) Then
                        MzMaxValue = M3(0)
                        MzMaxnode = ObjectName(j)
                        MzMaxComb = CombName(i)
                    End If

                    If MzMinValue > M3(0) Then
                        MzMinValue = M3(0)
                        MzMinNode = ObjectName(j)
                        MzMinComb = CombName(i)
                    End If

                    If K = 0 Then
                        PrintLine(FileIndex, AlignS(ObjectName(j), 0, 5) + AlignS(i + 1, 0, 5) + " " + CombName(i), Microsoft.VisualBasic.TAB(23), AlignS(F1(0)) + AlignS(F2(0)) + AlignS(F3(0)) + AlignS(M1(0)) + AlignS(M2(0)) + AlignS(M3(0)))
                    Else
                        PrintLine(FileIndex, AlignS(ObjectName(j), 0, 5) + AlignS(i + 1, 0, 5) + " " + CombName(i), Microsoft.VisualBasic.TAB(23), AlignS(F1(0)) + AlignS(F2(0)) + AlignS(F3(0)) + AlignS(M1(0)) + AlignS(M2(0)) + AlignS(M3(0)))
                    End If
pass:
                Next

            Next

            PrintLine(FileIndex, "  -------------------")
            PrintLine(FileIndex, "  Maximum & Minimum :")
            PrintLine(FileIndex, "  Quantity  Limit     Value    Unit       Node     Ldcmb")
            PrintLine(FileIndex, "  -----------------------------------------------------------------------------")
            PrintLine(FileIndex, "  FX        Max" + AlignS(FxMaxValue, 4, 13) + "   T      " + AlignS(FxMaxnode, 0, 7) + "      " + FxMaxComb.ToString)
            PrintLine(FileIndex, "            Min" + AlignS(FxMinValue, 4, 13) + "   T      " + AlignS(FxMinNode, 0, 7) + "      " + FxMinComb.ToString)
            PrintLine(FileIndex, "")
            PrintLine(FileIndex, "  FY        Max" + AlignS(FyMaxValue, 4, 13) + "   T      " + AlignS(FyMaxnode, 0, 7) + "      " + FyMaxComb.ToString)
            PrintLine(FileIndex, "            Min" + AlignS(FyMinValue, 4, 13) + "   T      " + AlignS(FyMinNode, 0, 7) + "      " + FyMinComb.ToString)
            PrintLine(FileIndex, "")
            PrintLine(FileIndex, "  FZ        Max" + AlignS(FzMaxValue, 4, 13) + "   T      " + AlignS(FzMaxnode, 0, 7) + "      " + FzMaxComb.ToString)
            PrintLine(FileIndex, "            Min" + AlignS(FzMinValue, 4, 13) + "   T      " + AlignS(FzMinNode, 0, 7) + "      " + FzMinComb.ToString)
            PrintLine(FileIndex, "  MX        Max" + AlignS(MxMaxValue, 4, 13) + "   T-M    " + AlignS(MxMaxnode, 0, 7) + "      " + FxMaxComb.ToString)
            PrintLine(FileIndex, "            Min" + AlignS(MxMinValue, 4, 13) + "   T-M    " + AlignS(MxMinNode, 0, 7) + "      " + FxMinComb.ToString)
            PrintLine(FileIndex, "")
            PrintLine(FileIndex, "  MY        Max" + AlignS(MyMaxValue, 4, 13) + "   T-M    " + AlignS(MyMaxnode, 0, 7) + "      " + FxMaxComb.ToString)
            PrintLine(FileIndex, "            Min" + AlignS(MyMinValue, 4, 13) + "   T-M    " + AlignS(MyMinNode, 0, 7) + "      " + FxMinComb.ToString)
            PrintLine(FileIndex, "")
            PrintLine(FileIndex, "  MZ        Max" + AlignS(MzMaxValue, 4, 13) + "   T-M    " + AlignS(MzMaxnode, 0, 7) + "      " + FxMaxComb.ToString)
            PrintLine(FileIndex, "            Min" + AlignS(MzMinValue, 4, 13) + "   T-M    " + AlignS(MzMinNode, 0, 7) + "      " + FxMinComb.ToString)
            PrintLine(FileIndex, "")

        Next

        MsgBox("Complete" & vbCrLf & SaveFileDia.FileName & vbCrLf & LraFileName, , "Support Reaction")
        FileClose(10)
        FileClose(20)


exitsub:



        ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Reaction Force")


    End Sub

    '反力檔輸出用Function
    Private Function AlignS(ByRef Num As String, Optional ByVal Dec As Integer = 3, Optional ByRef Space As Integer = 10) As String

        Dim Output As String
        Num = CDbl(Num)
        Output = FormatNumber(Num, Dec).PadLeft(Space)

        AlignS = Output.ToString
        Return AlignS

    End Function

    'RC桿件分類與定義Piece Mark
    Public hasELzero As Boolean = False

    Private Sub Classify_Section_Type(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        SapModel.SetPresentUnits(eUnits.Ton_cm_C)

        Dim ret As Long
        Dim NumberNames As Long
        Dim GroupName() As String

        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String

        Dim PropName As String
        Dim SAuto As String

        Dim FrameName() As String
        Dim Location() As Double
        Dim TopCombo() As String
        Dim TopArea() As Double
        Dim BotCombo() As String
        Dim BotArea() As Double
        Dim VmajorCombo() As String
        Dim VmajorArea() As Double
        Dim TLCombo() As String
        Dim TLArea() As Double
        Dim TTCombo() As String
        Dim TTArea() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String

        Dim IsLocked As Boolean = SapModel.GetModelIsLocked

        If IsLocked = False Then
            MsgBox("Run Analyze First", MsgBoxStyle.Information, "Member sort")
            GoTo endsub
        End If

        ret = SapModel.DesignConcrete.StartDesign

        'Dim RebarSize_Girder As String
        'Dim RebarSize_Beam As String
        'Dim RebarSize_Torsion As String

        'Dim RebarArea_Pri As Double
        'Dim RebarArea_Sec As Double
        'Dim RebarArea_Tor As Double
        'Dim RebarArea As Double

        'RebarSize_Girder = InputBox("Rebar Size - Girder " & vbCrLf & "EX : D25 / #8  etc...", "Input Rebar Size - Girder", "D25").ToUpper

        'RebarSize_Beam = InputBox("Rebar Size - Beam  " & vbCrLf & "EX : D25 / #8 etc...", "Input Rebar Size - Beam", "D25").ToUpper
        'RebarSize_Torsion = InputBox("Rebar Size - Torsion  " & vbCrLf & "EX : D19 / #6  etc...", "Input Rebar Size - Torsion", "D19").ToUpper

        'RebarArea_Pri = GetRebarArea(RebarSize_Girder)
        'If RebarArea_Pri = -1 Then MsgBox("Can't find " & RebarSize_Girder & " please contact RD Team")
        'RebarArea_Sec = GetRebarArea(RebarSize_Beam)
        'If RebarArea_Sec = -1 Then MsgBox("Can't find " & RebarSize_Beam & " please contact RD Team")
        'RebarArea_Tor = GetRebarArea(RebarSize_Torsion)
        'If RebarArea_Tor = -1 Then MsgBox("Can't find " & RebarSize_Torsion & " please contact RD Team")

        Dim frmSelectRebar As New Select_Rebar_Size
        frmSelectRebar.ShowDialog()

        If RebarArea_Pri = -1 Or RebarArea_Sec = -1 Or RebarArea_Tor = -1 Then
            GoTo EndSub
        End If


        Dim Point1, Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        '===================================
        ret = SapModel.SelectObj.All
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
        SapModel.SetPresentUnits(eUnits.Ton_cm_C)

        For i = 0 To NumberItems - 1
            If ObjectType(i) = 2 Then
                ret = SapModel.FrameObj.GetPoints(ObjectName(i), Point1, Point2)   'get point

                ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)     'point1 coord
                ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)     'point2 coord
                elevSort(Z1, Z2, Point1, Point2)
            End If

        Next

        If hasELzero = True Then Lcount = Lcount + 1

        ReDim Preserve elev(Lcount - 1)



        Array.Sort(elev)      'Array 為double type

        Dim Level(elev.Count - 1) As String

        For j = 0 To elev.Count - 1
            Level(j) = j + 1
        Next

        '====================================

        ret = SapModel.GroupDef.GetNameList(NumberNames, GroupName)

        Dim BeamDataCol(NumberNames, 12) As String       '只存桿件資料    是否需改用物件存資料?
        Dim iBeam As Integer = 1
        Dim iCol As Integer = 1
        Dim iMember As Integer = 1
        Dim iDelete As Integer

        Dim MyOption() As Integer
        Dim PMMCombo() As String
        Dim PMMArea() As Double
        Dim PMMRatio() As Double
        Dim AVmajor() As Double
        Dim VminorCombo() As String
        Dim AVminor() As Double

        SapModel.SetPresentUnits(eUnits.Ton_cm_C)
        For i = 0 To NumberNames - 1

            If IsNumeric(GroupName(i)) = True Then
                ret = SapModel.SelectObj.ClearSelection
                ret = SapModel.SelectObj.Group(GroupName(i))     'Revit ID 

                ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)   'Frame No.

                If NumberItems = 0 Then

                    ret = SapModel.GroupDef.Delete(GroupName(i))
                    iDelete += 1
                    GoTo skip
                ElseIf NumberItems > 1 Then
                    MsgBox("Group have more than 1 physical member" & vbCrLf & "Group  :    " & GroupName(i))
                    GoTo EndSub
                End If


                If ObjectType(0) = 2 Then
                    ret = SapModel.FrameObj.GetSection(ObjectName(0), PropName, SAuto)  'Section Name (Mark)

                    If PropName.Contains("RC") Then
                        BeamDataCol(iMember, 3) = "RC"
                        iMember += 1
                    Else
                        BeamDataCol(iMember, 3) = "STEEL"
                        iMember += 1
                    End If

                    If BeamDataCol(iMember - 1, 3) = "RC" Then
                        GoTo readdata
                    Else
                        GoTo skip
                    End If
readdata:

                    '==========以節點位置判斷桿件Type
                    ret = SapModel.FrameObj.GetPoints(ObjectName(0), Point1, Point2)   'get point

                    ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)     'point1 coord
                    ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)     'point2 coord

                    Dim SecType As String
                    SecType = MemberMark(X1, Y1, Z1, X2, Y2, Z2, Point1, Point2)

                    If SecType.Contains("C") Then

                        '待完成項目 - 20131107
                        '此處需處理SAP2000中，被判定為Column的桿件資料儲存程式碼
                        '要加ColumnDataCol 或是新建Column Data Structure
                        '讀取column 定義

                        '=================================================

                        ret = SapModel.DesignConcrete.GetSummaryResultsColumn(ObjectName(0), NumberItems, FrameName, _
                    MyOption, Location, PMMCombo, PMMArea, PMMRatio, VmajorCombo, AVmajor, VminorCombo, AVminor, ErrorSummary, WarningSummary)







                        '=================================================
                    ElseIf SecType.Contains("G") Or SecType.Contains("B") Then

                        ret = SapModel.DesignConcrete.GetSummaryResultsBeam(ObjectName(0), NumberItems, FrameName, Location _
               , TopCombo, TopArea, BotCombo, BotArea, VmajorCombo, VmajorArea, TLCombo, TLArea, TTCombo, TTArea _
               , ErrorSummary, WarningSummary)           'Rebar Area  (R1 ~ R7)  這API只能讀取Beam的結果
                        '=======================
                        If ret = 1 Then
                            MsgBox("Error when loading Member :  " & ObjectName(0))
                            GoTo skip
                        End If
                        '=======================
                    End If

                    '=============================================

                    If SecType.Contains("G") Then
                        RebarArea = RebarArea_Pri
                    ElseIf SecType.Contains("B") Then
                        RebarArea = RebarArea_Sec
                    ElseIf SecType.Contains("C") Then
                        iCol += 1
                        BeamDataCol(iMember - 1, 0) = GroupName(i)
                        BeamDataCol(iMember - 1, 1) = ObjectName(0)
                        BeamDataCol(iMember - 1, 2) = PropName
                        BeamDataCol(iMember - 1, 11) = SecType
                        GoTo skip
                    End If

                    BeamDataCol(iBeam, 0) = GroupName(i)
                    BeamDataCol(iBeam, 1) = ObjectName(0)
                    BeamDataCol(iBeam, 2) = PropName

                    '=============抓左中右最大值寫這裡 ， 同時換算成鋼筋支數
                    Dim SegArea As New Seg_RebarArea
                    SegArea = FindMaxRebarArea(TopArea, BotArea, TLArea, SegArea)

                    BeamDataCol(iBeam, 4) = Math.Ceiling(SegArea.TopLeft / RebarArea)
                    BeamDataCol(iBeam, 5) = Math.Ceiling(SegArea.BotLeft / RebarArea)

                    BeamDataCol(iBeam, 6) = Math.Ceiling(SegArea.TopMiddle / RebarArea)
                    BeamDataCol(iBeam, 7) = Math.Ceiling(SegArea.BotMiddle / RebarArea)

                    BeamDataCol(iBeam, 8) = Math.Ceiling(SegArea.TopRight / RebarArea)
                    BeamDataCol(iBeam, 9) = Math.Ceiling(SegArea.BotRight / RebarArea)

                    BeamDataCol(iBeam, 10) = Math.Ceiling(SegArea.Torsionbar / RebarArea_Tor)
                    '====================
                    '舊的
                    'BeamDataCol(iBeam, 4) = Math.Ceiling(TopArea(0) / RebarArea)
                    'BeamDataCol(iBeam, 5) = Math.Ceiling(BotArea(0) / RebarArea)

                    'BeamDataCol(iBeam, 6) = Math.Ceiling(TopArea(4) / RebarArea)
                    'BeamDataCol(iBeam, 7) = Math.Ceiling(BotArea(4) / RebarArea)

                    'BeamDataCol(iBeam, 8) = Math.Ceiling(TopArea(8) / RebarArea)
                    'BeamDataCol(iBeam, 9) = Math.Ceiling(BotArea(8) / RebarArea)

                    'BeamDataCol(iBeam, 10) = Math.Ceiling(TLArea(4) / RebarArea)

                    BeamDataCol(iBeam, 11) = SecType

                    iBeam += 1
                End If
skip:
            End If

        Next
        ret = SapModel.SelectObj.ClearSelection

        iCol = iCol - 1
        iBeam = iBeam - 1
        iMember = iMember - 1
        'Mark名稱初始化
        'BeamDataCol(0, 10) = BeamDataCol(0, 10) + "1"



        '==============此處開始========


        ''            Dim MarkCounter As Integer = 1
        ''            Dim memberCounter As Integer = 0
        ''            Dim TypeCounter(NumberNames) As Integer
        ''            TypeCounter(1) = 1

        ''            BeamDataCol(1, 11) = 1           'Type Counter 初始化

        ''            For i = 2 To NumberNames - 1
        ''                For j = 0 To memberCounter

        ''                    For k = 1 To i
        ''                        If BeamDataCol(i, 10) = BeamDataCol(k, 10) Then
        ''                            If BeamDataCol(i, 2) = BeamDataCol(k, 2) Then
        ''                                If CompareBeamRebar(BeamDataCol, i, k) = True Then
        ''                                    BeamDataCol(i, 11) = BeamDataCol(k, 11)
        ''                                    GoTo NextMember
        ''                                Else
        ''                                    MarkCounter += 1
        ''                                    BeamDataCol(i, 11) = MarkCounter
        ''                                    GoTo NextMember
        ''                                End If
        ''                            End If
        ''                        Else
        ''                            MarkCounter += 1
        ''                            BeamDataCol(i, 11) = MarkCounter
        ''                            GoTo NextMember
        ''                        End If
        ''                    Next




        ''                Next
        ''NextMember:     TypeCounter(i) = BeamDataCol(i, 11)

        ''            Next
        ''Dim MarkCounter(NumberNames - 1, 300) As String    '最多300種piecemark
        ''Dim counterA, counterB As Integer


        ''For i = 0 To NumberNames - 1
        ''    If BeamDataCol(i, 0) <> Nothing And BeamDataCol(i, 3) = "RC" Then
        ''        If MarkCounter.Exists(BeamDataCol(i, 11)) = False Then

        ''        End If




        ''    End If
        ''Next

        Dim MarkList As List(Of String) = New List(Of String)

        For i = 1 To iBeam
            If MarkList.Contains(BeamDataCol(i, 11)) = False Then
                MarkList.Add(BeamDataCol(i, 11))
            End If
        Next

        Dim GroupCount As Integer
        Dim TotalCount As Integer
        Dim MemGroup As List(Of Object) = New List(Of Object)
        Dim MemGroup_Name As List(Of String) = New List(Of String)
        Dim SameType As Boolean = False
        Dim MarkSerialNo As Integer

        For i = 0 To MarkList.Count - 1
            For j = 1 To iBeam
                If BeamDataCol(j, 11) = MarkList(i) Then
                    MemGroup.Add(j)     'Add Item Index to Container
                    MemGroup_Name.Add(BeamDataCol(j, 0))
                    If GroupCount = 0 Then BeamDataCol(j, 12) = "01"
                    GroupCount += 1
                    TotalCount += 1
                End If
            Next

            Dim counter1 As Integer = 0
            Dim counter2 As Integer = 2

            For k = 1 To MemGroup.Count - 1
                For m As Integer = 0 To counter1

                    If CompareBeamRebar(BeamDataCol, MemGroup(k), MemGroup(m)) = True Then
                        MarkSerialNo = BeamDataCol(MemGroup(m), 12)
                        BeamDataCol(MemGroup(k), 12) = MarkSerialNo.ToString("D2")
                        BeamDataCol(MemGroup(k), 12) = BeamDataCol(MemGroup(m), 12)
                        SameType = True
                    End If

                Next
                counter1 += 1
                If SameType = False Then
                    BeamDataCol(MemGroup(k), 12) = counter2.ToString("D2")
                    counter2 += 1
                End If
                SameType = False

            Next
            GroupCount = 0
            MemGroup.Clear()

        Next

        Dim Table(iBeam, 1) As String
        For i = 1 To iBeam
            Table(i - 1, 0) = BeamDataCol(i, 0)
            Table(i - 1, 1) = BeamDataCol(i, 11).ToString + BeamDataCol(i, 12).ToString

        Next

        '==========SAP 解鎖、新增斷面、指定斷面

        If MsgBox("程式將解鎖模型以建立新斷面", MsgBoxStyle.YesNo, "Member Sort") = MsgBoxResult.Yes Then
            ret = SapModel.SetModelIsLocked(False)
        Else
            MsgBox("End Program")
            GoTo EndSub
        End If

        '==============設定Output Station 為  17  (2016/11/23 修改 11 >> 17)
        ret = SapModel.SelectObj.All
        ret = SapModel.FrameObj.SetOutputStations("ALL", 2, 0, 17, True, True, eItemType.SelectedObjects)
        ret = SapModel.SelectObj.ClearSelection
        '=======================================
        'ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        'For i = 0 To NumberItems - 1
        '    If ObjectType(i) = 2 Then
        '        ret=SapModel.FrameObj .SetOutputStations (ObjectName(i),
        '    End If
        'Next


        'get rectangle
        Dim FileName As String
        Dim MatProp As String
        Dim t3 As Double
        Dim t2 As Double
        Dim Color As Long
        Dim Notes As String
        Dim GUID As String
        Dim NewSectionName As String
        Dim AreaTL, AreaTR, AreaBL, AreaBR As Double

        'get beam rebar
        Dim RebarName As String
        Dim MatPropLong As String
        Dim MatPropConfine As String
        Dim CoverTop As Double
        Dim CoverBot As Double
        Dim TopLeftArea As Double
        Dim TopRightArea As Double
        Dim BotLeftArea As Double
        Dim BotRightArea As Double

        SapModel.SetPresentUnits(eUnits.Ton_cm_C)

        For i = 1 To iBeam
            If BeamDataCol(i, 11).Contains("C") = False Then

                If BeamDataCol(i, 11).Contains("G") = True Then
                    AreaTL = BeamDataCol(i, 4) * RebarArea_Pri
                    AreaTR = BeamDataCol(i, 8) * RebarArea_Pri
                    AreaBL = BeamDataCol(i, 5) * RebarArea_Pri
                    AreaBR = BeamDataCol(i, 9) * RebarArea_Pri
                ElseIf BeamDataCol(i, 11).Contains("B") = True Then
                    AreaTL = BeamDataCol(i, 4) * RebarArea_Sec
                    AreaTR = BeamDataCol(i, 8) * RebarArea_Sec
                    AreaBL = BeamDataCol(i, 5) * RebarArea_Sec
                    AreaBR = BeamDataCol(i, 9) * RebarArea_Sec
                End If

                ret = SapModel.PropFrame.GetRectangle(BeamDataCol(i, 2), FileName, MatProp, t3, t2, Color, Notes, GUID)
                NewSectionName = BeamDataCol(i, 11).ToString + BeamDataCol(i, 12).ToString + "-" + BeamDataCol(i, 2).ToString

                If BeamDataCol(i, 2).Contains(NewSectionName) Then
                    MsgBox("桿件已分類，請輸出結果Excel檔 ， 程式結束" & vbCrLf & "")
                    GoTo endsub
                End If

                '===========建立新名稱桿件與紀錄主筋Size
                If BeamDataCol(i, 11).Contains("G") Then
                    ret = SapModel.PropFrame.SetRectangle(NewSectionName, MatProp, t3, t2, -1, RebarSize_Girder)
                ElseIf BeamDataCol(i, 11).Contains("B") Then
                    ret = SapModel.PropFrame.SetRectangle(NewSectionName, MatProp, t3, t2, -1, RebarSize_Beam)
                End If

                'ret = SapModel.SelectObj.ClearSelection
                ret = SapModel.PropFrame.GetRebarBeam(BeamDataCol(i, 2), MatPropLong, MatPropConfine, CoverTop, CoverBot, TopLeftArea, TopRightArea, BotLeftArea, BotRightArea)


                '2015/10/08 修改為先不要寫入主筋量
                'ret = SapModel.PropFrame.SetRebarBeam(NewSectionName, MatPropLong, MatPropConfine, CoverTop, CoverBot, AreaTL, AreaTR, AreaBL, AreaBR)
                ret = SapModel.PropFrame.SetRebarBeam(NewSectionName, MatPropLong, MatPropConfine, CoverTop, CoverBot, 0, 0, 0, 0)

                ret = SapModel.FrameObj.SetSection(BeamDataCol(i, 1), NewSectionName)

            End If
        Next

        MsgBox("Finish", , "Classify Members")
EndSub:
        ISapPlugin.Finish(0)

    End Sub

    'Rebar Data Structure(Beam)
    Public Structure Seg_RebarArea
        Public TopLeft As Double
        Public TopMiddle As Double
        Public TopRight As Double
        Public BotLeft As Double
        Public BotMiddle As Double
        Public BotRight As Double
        Public Torsionbar As Double
    End Structure

    '找出各段最大鋼筋量,此處資料尚未限定station數
    Public Function FindMaxRebarArea(ByVal TopArea() As Double, ByVal BotArea() As Double, ByVal TorRebar() As Double, ByVal Rebar As Seg_RebarArea)
        Dim seg As Integer = Math.Floor(TopArea.Count / 3)
        Dim T_leftMax, T_midMax, T_rightMax As Double
        Dim B_leftMax, B_midMax, B_rightMax As Double
        Dim TorsionRebarMax As Double

        For i = 1 To TopArea.Count
            If i <= seg And T_leftMax < TopArea(i - 1) Then
                T_leftMax = TopArea(i - 1)
            End If

            If i > TopArea.Count - seg And T_rightMax < TopArea(i - 1) Then
                T_rightMax = TopArea(i - 1)
            End If

            If i > seg And i <= TopArea.Count - seg And T_midMax < TopArea(i - 1) Then
                T_midMax = TopArea(i - 1)
            End If
        Next

        Rebar.TopLeft = T_leftMax
        Rebar.TopMiddle = T_midMax
        Rebar.TopRight = T_rightMax

        For i = 1 To BotArea.Count
            If i <= seg And B_leftMax < BotArea(i - 1) Then
                B_leftMax = BotArea(i - 1)
            End If

            If i > BotArea.Count - seg And B_rightMax < BotArea(i - 1) Then
                B_rightMax = BotArea(i - 1)
            End If

            If i > seg And i <= BotArea.Count - seg And B_midMax < BotArea(i - 1) Then
                B_midMax = BotArea(i - 1)
            End If
        Next

        For i = 1 To TopArea.Count
            If TorsionRebarMax < TorRebar(i - 1) Then
                TorsionRebarMax = TorRebar(i - 1)
            End If

        Next
        Rebar.Torsionbar = TorsionRebarMax

        Rebar.BotLeft = B_leftMax
        Rebar.BotMiddle = B_midMax
        Rebar.BotRight = B_rightMax

        Return Rebar
    End Function

    '斷面鋼筋量比較
    Private Function CompareBeamRebar(ByRef BeamDataCol As Array, ByRef Member1 As Integer, ByRef Member2 As Integer) As Boolean

        Dim Temp(7, 1) As String
        Temp(0, 0) = BeamDataCol(Member1, 2)
        Temp(0, 1) = BeamDataCol(Member2, 2)
        Temp(1, 0) = BeamDataCol(Member1, 4)
        Temp(1, 1) = BeamDataCol(Member2, 4)
        Temp(2, 0) = BeamDataCol(Member1, 5)
        Temp(2, 1) = BeamDataCol(Member2, 5)
        Temp(3, 0) = BeamDataCol(Member1, 6)
        Temp(3, 1) = BeamDataCol(Member2, 6)
        Temp(4, 0) = BeamDataCol(Member1, 7)
        Temp(4, 1) = BeamDataCol(Member2, 7)
        Temp(5, 0) = BeamDataCol(Member1, 8)
        Temp(5, 1) = BeamDataCol(Member2, 8)
        Temp(6, 0) = BeamDataCol(Member1, 9)
        Temp(6, 1) = BeamDataCol(Member2, 9)
        Temp(7, 0) = BeamDataCol(Member1, 1)
        Temp(7, 1) = BeamDataCol(Member2, 1)

        If BeamDataCol(Member1, 2) = BeamDataCol(Member2, 2) And BeamDataCol(Member1, 4) = BeamDataCol(Member2, 4) And BeamDataCol(Member1, 5) = BeamDataCol(Member2, 5) And BeamDataCol(Member1, 6) = BeamDataCol(Member2, 6) And BeamDataCol(Member1, 7) = BeamDataCol(Member2, 7) And BeamDataCol(Member1, 8) = BeamDataCol(Member2, 8) And BeamDataCol(Member1, 9) = BeamDataCol(Member2, 9) Then
            Return True
        End If

    End Function

    '自動定出Piece Mark
    Private Function MemberMark(ByRef Sx As Double, ByRef Sy As Double, ByRef Sz As Double, ByRef Ex As Double, ByRef Ey As Double, ByRef Ez As Double, ByRef P1 As String, ByRef P2 As String) As String

        '輸入起點/終點座標 透過規則定出該桿件類別

        Dim Lev As String

        If Math.Abs(Sz - Ez) < 0.001 Then
            For i = 0 To Lcount - 1
                If Math.Abs(Sz - elev(i)) < 0.001 Then
                    Lev = (i + 1).ToString
                    Exit For
                End If
            Next
        Else
            For i = 0 To Lcount - 1
                If Ez = elev(i) Then
                    Lev = (i + 1).ToString
                    Exit For
                End If
            Next
        End If

        Dim BeamGirderDetP1 As Boolean = False
        Dim BeamGirderDetP2 As Boolean = False

        Dim BeamGirderDet As Boolean = False

        If ConnectiveJoint Is Nothing Then
            ReDim ConnectiveJoint(1)
        End If


        '判斷是大梁還是小梁============================
        '檢查第一點有無和柱相接
        For i = 0 To ConnectiveJoint.Count - 2
            If P1 = ConnectiveJoint(i) Then
                BeamGirderDetP1 = True
                Exit For
            End If
        Next

        '檢查第二點有無和柱相接
        For i = 0 To ConnectiveJoint.Count - 2
            If P2 = ConnectiveJoint(i) Then
                BeamGirderDetP2 = True
                Exit For
            End If
        Next

        '兩點皆和柱相接的為大梁
        If BeamGirderDetP1 = True And BeamGirderDetP2 = True Then
            BeamGirderDet = True
        End If
        '============================


        If BeamGirderDet = True And Math.Abs(Sz - Ez) < 0.001 Then
            Lev = Lev + "G"
        ElseIf BeamGirderDet = False And Math.Abs(Sz - Ez) < 0.001 Then
            Lev = Lev + "B"
        End If


        If Math.Abs(Sz - Ez) >= 0.01 And Sx = Ex And Sy = Ey Then
            Lev = Lev + "C"
        Else
            If Math.Round(Sy, 2) = Math.Round(Ey, 2) Then Lev = Lev + "X"
            If Math.Round(Sx, 2) = Math.Round(Ex, 2) Then Lev = Lev + "Y"

        End If

        Return Lev

    End Function

    '找出所有的elevation
    Public Sub elevSort(ByRef Z1 As Double, ByRef Z2 As Double, ByRef P1 As String, ByRef P2 As String)

        Dim Repeat As Boolean = False


        If Math.Abs(Z1 - Z2) < 0.001 Then
            ReDim Preserve elev(Lcount)
            For i = 0 To Lcount
                If Math.Abs(Z1 - elev(i)) < 0.001 Then
                    Repeat = True
                End If
            Next

            If Z1 = 0 And Z2 = 0 Then
                HasELzero = True
            End If


            If Repeat = False Then
                elev(Lcount) = Z1
                Lcount += 1
            End If
        End If




        If Math.Abs(Z1 - Z2) >= 0.001 Then
            ConnCount += 2
            ReDim Preserve ConnectiveJoint(ConnCount)
            ConnectiveJoint(ConnCount - 2) = CInt(P1)
            ConnectiveJoint(ConnCount - 1) = CInt(P2)
        End If


    End Sub

    '輸出RC設計結果
    Private Sub Output_RC_Design_Result(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim ReinBemInf As New OpenFileDialog

        ReinBemInf.Filter = "Reinbem.inf|*.inf"
        ReinBemInf.Title = "Select Beam Rebar Setting Data"


        If ReinBemInf.ShowDialog() <> DialogResult.OK Then
            MsgBox("Can't Load Beam Rebar Information File")
            GoTo endsub2
        End If

        Dim charseparators() As Char = " "
        Dim Dummy As String
        Dim Line1() As String
        Dim Line2() As String
        Dim temp() As String
        Dim RebarTypes As Integer
        Dim WidthTypes As Integer
        'Dim WidthList() As Integer

        FileOpen(10, ReinBemInf.FileName, OpenMode.Input)
        'Dummy = LineInput(10)

        Line1 = LineInput(10).Split(charseparators, StringSplitOptions.RemoveEmptyEntries)
        RebarTypes = Line1(0)
        WidthTypes = Line1(1)

        Line2 = LineInput(10).Split(charseparators, StringSplitOptions.RemoveEmptyEntries)

        If WidthTypes <> Line2.Count - 4 Then
            MsgBox("Check *.inf File" & vbCrLf & "Data Number Inconsistency")
            GoTo endsub2
        End If


        Dim RebarData(RebarTypes - 1, WidthTypes) As String
        Dim ii, j As Integer

        Do Until EOF(10)
            temp = LineInput(10).Split(charseparators, StringSplitOptions.RemoveEmptyEntries)
            RebarData(ii, 0) = temp(0)
            For j = 1 To WidthTypes
                RebarData(ii, j) = temp(j + 3)
            Next
            ii = ii + 1
            If ii = RebarTypes + 1 Then
                MsgBox("檢查*.inf 第一行")
                GoTo endsub2
            End If

            j = 0
        Loop

        FileClose(10)


        SapModel.SetPresentUnits(eUnits.kgf_cm_C)

        Dim ret As Long
        Dim NumberNames As Long
        Dim GroupName() As String

        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String

        Dim PropName As String
        Dim SAuto As String

        Dim FrameName() As String
        Dim Location() As Double
        Dim TopCombo() As String
        Dim TopArea() As Double
        Dim BotCombo() As String
        Dim BotArea() As Double
        Dim VmajorCombo() As String
        Dim VmajorArea() As Double
        Dim TLCombo() As String
        Dim TLArea() As Double
        Dim TTCombo() As String
        Dim TTArea() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String

        Dim fc As Double
        Dim IsLightweight As Boolean
        Dim fcsfactor As Double
        Dim SSType As Long
        Dim SSHysType As Long
        Dim StrainAtfc As Double
        Dim StrainUltimate As Double
        Dim FinalSlope As Double
        Dim FrictionAngle As Double
        Dim DilatationalAngle As Double

        Dim Fy As Double
        Dim Fu As Double
        Dim eFy As Double
        Dim eFu As Double
        'Dim SSType As Long
        'Dim SSHysType As Long
        Dim StrainAtHardening As Double
        'Dim StrainUltimate As Double
        Dim UseCaltransSSDefaults As Boolean


        'Dim RebarArea_Pri As Double
        'Dim RebarArea_Sec As Double
        'Dim RebarArea_Tor As Double
        'Dim RebarArea As Double

        'RebarArea_Pri = InputBox("Rebar Area - Girder  (cm^2)" & vbCrLf & "EX : D25 = 5.07" & vbCrLf & "#10 = 8.14" & vbCrLf & "etc...")
        'RebarArea_Sec = InputBox("Rebar Area - Beam    (cm^2)" & vbCrLf & "EX : D25 = 5.07" & vbCrLf & "#10 = 8.14" & vbCrLf & "etc...")
        'RebarArea_Tor = InputBox("Rebar Area - Torsion    (cm^2)" & vbCrLf & "EX : D19 = 2.87" & vbCrLf & "#10 = 8.14" & vbCrLf & "etc...")

        'Dim Point1, Point2 As String
        'Dim X1, Y1, Z1 As Double
        'Dim X2, Y2, Z2 As Double

        '================Output Sheet 2 - RC member


        ret = SapModel.GroupDef.GetNameList(NumberNames, GroupName)

        Dim FileName As String
        Dim MatProp As String
        Dim t3 As Double
        Dim t2 As Double
        Dim Color As Long
        Dim Notes As String
        Dim GUID As String
        Dim NewSectionName As String
        Dim AreaTL, AreaTR, AreaBL, AreaBR As Double

        Dim RebarName As String
        Dim MatPropLong As String
        Dim MatPropConfine As String
        Dim CoverTop As Double
        Dim CoverBot As Double
        Dim TopLeftArea As Double
        Dim TopRightArea As Double
        Dim BotLeftArea As Double
        Dim BotRightArea As Double

        Dim Ang As Double
        Dim Advanced As Boolean

        Dim BeamRebarLimit As Integer

        Dim Sheet_RC_Member(NumberNames, 2) As String

        If SapModel.GetModelIsLocked = False Then
            MsgBox("Run Analyze First")
            GoTo endsub2
        End If

        Dim StirrupSize As String
        Dim SideBarSize As String
        StirrupSize = InputBox("輸入箍筋尺寸   EX :  D10 / #3", "Output RC Design Result", "D13").ToUpper
        SideBarSize = InputBox("輸入腹筋尺寸   EX :  D19 / #6", "Output RC Design Result", "D19").ToUpper

        Dim Physical(NumberNames) As MemberProp

        For i = 0 To NumberNames - 1
            Dim Member As New MemberProp

            If IsNumeric(GroupName(i)) = True Then

                Member.Stirrup = StirrupSize
                Member.SideBar = SideBarSize

                Member.OID = GroupName(i)

                ret = SapModel.SelectObj.ClearSelection
                ret = SapModel.SelectObj.Group(GroupName(i))     'Revit ID 



                ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)   'Frame No.
                If ObjectType(0) = 2 Then

                    ret = SapModel.FrameObj.GetSection(ObjectName(0), PropName, SAuto)  'Section Name (Size & Mark)

                    '==========以名稱判斷是RC或steel
                    If PropName.Contains("RC") = False Then
                        GoTo readnextmember
                    End If

                    '=========================


                    GetSizeMark(Member, PropName)
                    Sheet_RC_Member(i, 0) = GroupName(i)
                    Sheet_RC_Member(i, 1) = Member.FullName
                    Sheet_RC_Member(i, 2) = Member.PieceMark

                    '=====================================================
                    If PropName.Contains("G") Or PropName.Contains("B") Then
                        Member.Type = "Beam"
                    ElseIf PropName.Contains("C") Then
                        Member.Type = "Column"
                    Else
                        Member.Type = "Unknow"
                    End If

                    ret = SapModel.PropFrame.GetRectangle(PropName, FileName, MatProp, t3, t2, Color, Notes, GUID)
                    Member.MainBar = Notes              '在Note中存放主筋Size
                    Member.ConcreteType = MatProp       '混凝土材料名稱

                    SapModel.SetPresentUnits(eUnits.kgf_cm_C)
                    ret = SapModel.PropMaterial.GetOConcrete_1(MatProp, fc, IsLightweight, fcsfactor, SSType, SSHysType, StrainAtfc, StrainUltimate, FinalSlope, FrictionAngle, DilatationalAngle)


                    Member.ConcreteStrength = fc        'Fc (kgf/cm^2)
                    Member.Width = Math.Round(t2, 0)    '從SAP讀取斷面尺寸
                    Member.Depth = Math.Round(t3, 0)



                    If Member.Type = "Beam" Then

                        '讀取使用者之前設計輸入之端部主筋量(不含中間段)
                        ret = SapModel.PropFrame.GetRebarBeam(PropName, MatPropLong, MatPropConfine, CoverTop, CoverBot, TopLeftArea, TopRightArea, BotLeftArea, BotRightArea)
                        Member.MainBarMat = MatPropLong
                        ret = SapModel.PropMaterial.GetORebar(MatPropLong, Fy, Fu, eFy, eFu, SSType, SSHysType, StrainAtHardening, StrainUltimate, UseCaltransSSDefaults)
                        Member.MainBarFy = Fy
                        Member.SideBarFy = Fy

                        Member.StirrupMat = MatPropConfine
                        ret = SapModel.PropMaterial.GetORebar(MatPropConfine, Fy, Fu, eFy, eFu, SSType, SSHysType, StrainAtHardening, StrainUltimate, UseCaltransSSDefaults)
                        Member.StirrupFy = Fy

                        Member.CoverTop = Math.Round(CoverTop, 3)
                        Member.CoverBot = Math.Round(CoverBot, 3)
                        Member.top1_S = Math.Round(TopLeftArea / GetRebarArea(Member.MainBar))
                        Member.top1_E = Math.Round(TopRightArea / GetRebarArea(Member.MainBar))
                        Member.bot1_S = Math.Round(BotLeftArea / GetRebarArea(Member.MainBar))
                        Member.bot1_E = Math.Round(BotRightArea / GetRebarArea(Member.MainBar))

                        '使用inf檔之鋼筋數量限制
                        BeamRebarLimit = 0
                        For j = 4 To WidthTypes + 3
                            If Member.Width = Line2(j) Then
                                For k = 0 To RebarTypes - 1
                                    If Member.MainBar = RebarData(k, 0) Then
                                        BeamRebarLimit = RebarData(k, j - 3)
                                        Exit For
                                    End If
                                Next
                            End If
                        Next

                        If BeamRebarLimit = 0 Then
                            MsgBox("Can't Find Limit in *.inf")
                            MsgBox("Bar Size :  " & Member.MainBar)
                            GoTo endsub2

                        End If

                        '桿件旋轉角
                        ret = SapModel.FrameObj.GetLocalAxes(ObjectName(0), Ang, Advanced)
                        Member.RotationAng = Ang

                        '只使用SAP分析的中段主筋、剪力箍筋、雙向扭力鋼筋，頭尾端部鋼筋使用預先設計之量
                        ret = SapModel.DesignConcrete.GetSummaryResultsBeam(ObjectName(0), NumberItems, FrameName, Location _
                      , TopCombo, TopArea, BotCombo, BotArea, VmajorCombo, VmajorArea, TLCombo, TLArea, TTCombo, TTArea _
                      , ErrorSummary, WarningSummary)           'Rebar Area  (R1 ~ R7)  這API有問題，讀Column ret仍然回傳0，這裡已修改為只讀"Beam"

                        '搜尋中間段之最大鋼筋量
                        Dim TopMiddleRebarMax As Double = 0
                        Dim BottomMiddlerebarMax As Double = 0
                        For j = 3 To 7
                            If TopMiddleRebarMax < TopArea(j) Then TopMiddleRebarMax = TopArea(j)
                            If BottomMiddlerebarMax < BotArea(j) Then BottomMiddlerebarMax = BotArea(j)
                        Next
                        Member.top1_M = Math.Ceiling(TopMiddleRebarMax / GetRebarArea(Member.MainBar))
                        Member.bot1_M = Math.Ceiling(BottomMiddlerebarMax / GetRebarArea(Member.MainBar))

                        '計算 member.top1 / top2 / top3 & bot1 / bot2 / bot3 / 中段鋼筋
                        CalculateRebar(Member, BeamRebarLimit)

                        'If Member.OverRebar = True Then
                        '    MsgBox("Member Design Rebar Over Maximum Allowable Value!")
                        'End If

                        '開始箍筋計算


                        'Dim ShearBarArea(10) As Double
                        'Dim TorsionBarArea(10) As Double
                        'For j = 0 To 10
                        '    ShearBarArea(j) = VmajorArea(j) + 2 * TTArea(j)

                        'Next

                        'Dim StartStirrupMax As Double
                        'Dim MiddleStirrupMax As Double
                        'Dim EndStirrupMax As Double

                        For j = 0 To 2
                            If Member.Stirrup_VRebar_S < VmajorArea(j) Then Member.Stirrup_VRebar_S = VmajorArea(j)
                            If Member.Stirrup_TTrnRebar_S < TTArea(j) Then Member.Stirrup_TTrnRebar_S = TTArea(j)
                        Next

                        For j = 3 To 7
                            If Member.Stirrup_VRebar_M < VmajorArea(j) Then Member.Stirrup_VRebar_M = VmajorArea(j)
                            If Member.Stirrup_TTrnRebar_M < TTArea(j) Then Member.Stirrup_TTrnRebar_M = TTArea(j)
                        Next

                        For j = 8 To 10
                            If Member.Stirrup_VRebar_E < VmajorArea(j) Then Member.Stirrup_VRebar_E = VmajorArea(j)
                            If Member.Stirrup_TTrnRebar_E < TTArea(j) Then Member.Stirrup_TTrnRebar_E = TTArea(j)
                        Next

                        '換算成間距與支數
                        CalculateStirrup(Member, Member.Stirrup)

                        'Member.Stirrup_Num_S = Math.Ceiling(Member.Stirrup_Area_S * Member.Stirrup_Spacing_S / GetRebarArea(Member.Stirrup))
                        'Member.Stirrup_Num_M = Math.Ceiling(Member.Stirrup_Area_M * Member.Stirrup_Spacing_M / GetRebarArea(Member.Stirrup))
                        'Member.Stirrup_Num_E = Math.Ceiling(Member.Stirrup_Area_E * Member.Stirrup_Spacing_E / GetRebarArea(Member.Stirrup))

                        '資料儲存


                    End If

                    '=====================================================
                    Physical(i) = Member

                End If
            End If
readnextmember:
        Next

        'Call Excel API
        '=================================================================================
        '結果輸出至Excel

        On Error Resume Next
        '#一部電腦僅執行一個Excel Application, 就算中突開啟Excel也不會影響程式執行   
        '#在工作管理員中只會看見一個EXCEL.exe在執行，不會浪費電腦資源      
        '#引用正在執行的Excel Application   
        'xlApp = GetObject(, "Excel.Application")
        '#若發生錯誤表示電腦沒有Excel正在執行，需重新建立一個新的應用程式    
        'If Err.Number() <> 0 Then
        '    Err.Clear()
        '#執行一個新的Excel Application        
        xlApp = CreateObject("Excel.Application")
        'If Err.Number() <> 0 Then
        '    MsgBox("電腦沒有安裝Excel")
        '    GoTo endsub
        'End If
        'End If


        Dim OpenExcelDialog As New OpenFileDialog

        OpenExcelDialog.Filter = "Excel2007 File|*.xlsx"
        OpenExcelDialog.Title = "Select Initial Beam Rebar Data"
        If OpenExcelDialog.ShowDialog() = DialogResult.OK Then
            '打開已經存在的EXCEL工件簿文件
            xlBook = xlApp.Workbooks.Open(OpenExcelDialog.FileName)
        Else
            MsgBox("Can't Load File")
            GoTo endsub2
        End If

        '停用警告訊息        
        xlApp.DisplayAlerts = False
        '設置EXCEL對象可見        
        xlApp.Visible = True
        '設定活頁簿為焦點        
        xlBook.Activate()
        '顯示第一個子視窗       
        'xlBook.Parent.Windows(1).Visible = True
        '引用第一個工作表     
        'xlSheet = xlBook.Worksheets(1)
        '設定工作表為焦點     
        'xlSheet.Activate()

        xlSheet = xlBook.Worksheets(2)
        xlSheet.Activate()
        '================================================================================

        '================Output Sheet 2 - RC member
        '要加入Excel 儲存格格式變更指令 ==>  文字

        For i = 1 To NumberNames - 1
            xlSheet.Cells(i + 1, 1) = Sheet_RC_Member(i, 0)           'Member ID
            xlSheet.Cells(i + 1, 2) = Sheet_RC_Member(i, 1)           'Size(Section Name)
            xlSheet.Cells(i + 1, 3) = Sheet_RC_Member(i, 2)           'Piece Mark
        Next

        '================Output Sheet 3 - RC Beam


        xlSheet = xlBook.Worksheets(3)
        xlSheet.Activate()

        Dim RCBeam_i As Integer = 0
        Dim PieceMarkRec(NumberNames) As String
        Dim PM_Rec As Integer
        Dim PM_RepeatFlag As Boolean = False



        For i = 1 To NumberNames - 1

            If Physical(i).Type = "Beam" Then

                '=======檢查同PieceMark 資料是否已重複，若為第一次出現則輸出至Excel
                'For j = 0 To PieceMarkRec.Count - 1
                '    If Physical(i).PieceMark = PieceMarkRec(j) Then
                '        PM_RepeatFlag = True
                '        Exit For
                '    End If
                'Next


                If PieceMarkRec.Contains(Physical(i).PieceMark) Then
                    PM_RepeatFlag = True
                Else
                    RCBeam_i += 1    'Excel 數格子用，資料不重複cunter就加1
                End If

                PieceMarkRec(i) = Physical(i).PieceMark

                '==================================================================

                If PM_RepeatFlag = False Then

                    'With xlBook.Worksheets(3)
                    '    .Range(.Cel1s(RCBeam_i, 1), .Cells(RCBeam_i + 10, 7)).copy()
                    '    .Range("A12").select()

                    'End With

                    xlSheet.Range(xlApp.Cells((RCBeam_i - 1) * 11 + 1, 1), xlApp.Cells(RCBeam_i * 11, 7)).Copy(Destination:=xlApp.Worksheets(3).cells(RCBeam_i * 11 + 1, 1))


                    '輸出資料==============
                    If RCBeam_i = 1 Then
                        xlSheet.Cells(1, 1) = "UNIT"
                        xlSheet.Cells(1, 2) = "mm"
                    End If
                    '======================
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 2, 2) = Physical(i).Type
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 3, 2) = Physical(i).PieceMark
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 3, 4) = Physical(i).FullName
                    '======================
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 4, 2) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 4, 3) = Physical(i).top1_S
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 4, 4) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 4, 5) = Physical(i).top1_M
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 4, 6) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 4, 7) = Physical(i).top1_E
                    '======================
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 5, 2) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 5, 3) = Physical(i).top2_S
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 5, 4) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 5, 5) = Physical(i).top2_M
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 5, 6) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 5, 7) = Physical(i).top2_E
                    '======================
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 6, 2) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 6, 3) = Physical(i).top3_S
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 6, 4) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 6, 5) = Physical(i).top3_M
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 6, 6) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 6, 7) = Physical(i).top3_E
                    '======================
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 7, 2) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 7, 3) = Physical(i).bot3_S
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 7, 4) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 7, 5) = Physical(i).bot3_M
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 7, 6) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 7, 7) = Physical(i).bot3_E
                    '======================
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 8, 2) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 8, 3) = Physical(i).bot2_S
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 8, 4) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 8, 5) = Physical(i).bot2_M
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 8, 6) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 8, 7) = Physical(i).bot2_E
                    '======================
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 9, 2) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 9, 3) = Physical(i).bot1_S
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 9, 4) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 9, 5) = Physical(i).bot1_M
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 9, 6) = Physical(i).MainBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 9, 7) = Physical(i).bot1_E
                    '======================
                    'cm 換 mm
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 10, 2) = Physical(i).Stirrup_Num_S.ToString + "-" + Physical(i).Stirrup
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 10, 3) = Physical(i).Stirrup_Spacing_S * 10
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 10, 4) = Physical(i).Stirrup_Num_S.ToString + "-" + Physical(i).Stirrup
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 10, 5) = Physical(i).Stirrup_Spacing_M * 10
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 10, 6) = Physical(i).Stirrup_Num_S.ToString + "-" + Physical(i).Stirrup
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 10, 7) = Physical(i).Stirrup_Spacing_E * 10
                    '======================
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 11, 2) = Physical(i).SideBar
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 11, 3) = Physical(i).SideBar_Num
                    xlSheet.Cells((RCBeam_i - 1) * 11 + 11, 10) = Physical(i).WarningMsg
                End If
                'RCBeam_i += 1     'Counter 固定+1
                PM_RepeatFlag = False   ' Flag 復原
            End If
        Next


        MsgBox("END")
endsub2:
        ISapPlugin.Finish(0)
    End Sub

    'RC Member Data Structure
    Public Structure MemberProp
        Public OID As String
        Public FullName As String
        Public PieceMark As String
        Public Width As Double
        Public Depth As Double
        Public ConcreteType As String        '材料名稱
        Public ConcreteStrength As String    '材料強度  (kgf/cm^2)
        Public CoverTop, CoverBot As Double  '上下保護層厚度 邊緣到縱向筋中心   (cm)
        Public Type As String                '桿件類別

        '主筋種類 Dxx  /  各處鋼筋量 >> 支數
        Public MainBar As String
        Public MainBarMat As String
        Public MainBarFy As String
        Public top1_S, top2_S, top3_S As String
        Public top1_M, top2_M, top3_M As String
        Public top1_E, top2_E, top3_E As String
        Public bot1_S, bot2_S, bot3_S As String
        Public bot1_M, bot2_M, bot3_M As String
        Public bot1_E, bot2_E, bot3_E As String
        '腹筋種類 Dxx  /  Fy (kgf/cm^2)  /  腹筋面積  /  腹筋支數
        Public SideBar As String
        Public SideBarMat As String
        Public SideBarFy As String
        Public SideBar_Area As Double
        Public SideBar_Num As String
        '箍筋種類 Dxx  /  Fy (kgf/cm^2)  /  面積  /  支數  /  間距
        Public Stirrup As String
        Public StirrupMat As String
        Public StirrupFy As String
        Public Stirrup_VRebar_S, Stirrup_VRebar_M, Stirrup_VRebar_E As Double
        Public Stirrup_TTrnRebar_S, Stirrup_TTrnRebar_M, Stirrup_TTrnRebar_E As Double
        Public Stirrup_Num_S, Stirrup_Num_M, Stirrup_Num_E As Double
        Public Stirrup_Spacing_S, Stirrup_Spacing_M, Stirrup_Spacing_E As Double
        '桿件旋轉角(Degree)  /  鋼筋過量標記  /  警告訊息
        Public RotationAng As Double
        Public OverRebar As Boolean
        Public WarningMsg As String

    End Structure

    '計算各處鋼筋支數
    Public Function CalculateRebar(ByRef member As MemberProp, ByRef RebarLimit As Integer) As MemberProp

        Dim Total_Top_Start As Integer = member.top1_S
        Dim Total_Top_End As Integer = member.top1_E
        Dim Total_Bottom_Start As Integer = member.bot1_S
        Dim Total_Bottom_End As Integer = member.bot1_E

        Dim Total_Top_Middle As Integer = member.top1_M
        Dim Total_Bottom_Middle As Integer = member.bot1_M

        If Total_Top_Start / RebarLimit <= 1 Then
            member.top1_S = Total_Top_Start
            member.top2_S = 0
            member.top3_S = 0
        ElseIf Total_Top_Start / RebarLimit <= 2 Then
            member.top1_S = RebarLimit
            member.top2_S = Total_Top_Start - RebarLimit
            member.top3_S = 0
        ElseIf Total_Top_Start / RebarLimit <= 3 Then
            member.top1_S = RebarLimit
            member.top2_S = RebarLimit
            member.top3_S = Total_Top_Start - RebarLimit * 2
        Else
            'MsgBox("Over Maximum Allowable Value" & vbCrLf & "Member ID :  " & vbCrLf & member.OID)
            member.top1_S = "Over"
            member.top2_S = "Over"
            member.top3_S = "Over"
            member.OverRebar = True
            'GoTo endFunc
        End If


        If Total_Top_End / RebarLimit <= 1 Then
            member.top1_E = Total_Top_End
            member.top2_E = 0
            member.top3_E = 0
        ElseIf Total_Top_End / RebarLimit <= 2 Then
            member.top1_E = RebarLimit
            member.top2_E = Total_Top_End - RebarLimit
            member.top3_E = 0
        ElseIf Total_Top_End / RebarLimit <= 3 Then
            member.top1_E = RebarLimit
            member.top2_E = RebarLimit
            member.top3_E = Total_Top_End - RebarLimit * 2
        Else
            'MsgBox("Over Maximum Allowable Value" & vbCrLf & "Member ID :  " & vbCrLf & member.OID)
            member.top1_E = "Over"
            member.top2_E = "Over"
            member.top3_E = "Over"
            member.OverRebar = True
            'GoTo endFunc
        End If

        '=============
        If Total_Bottom_Start / RebarLimit <= 1 Then
            member.bot1_S = Total_Bottom_Start
            member.bot2_S = 0
            member.bot3_S = 0
        ElseIf Total_Bottom_Start / RebarLimit <= 2 Then
            member.bot1_S = RebarLimit
            member.bot2_S = Total_Bottom_Start - RebarLimit
            member.bot3_S = 0
        ElseIf Total_Bottom_Start / RebarLimit <= 3 Then
            member.bot1_S = RebarLimit
            member.bot2_S = RebarLimit
            member.bot3_S = Total_Bottom_Start - RebarLimit * 2
        Else
            'MsgBox("Over Maximum Allowable Value" & vbCrLf & "Member ID :  " & vbCrLf & member.OID)
            member.bot1_S = "Over"
            member.bot2_S = "Over"
            member.bot3_S = "Over"
            member.OverRebar = True
            'GoTo endFunc
        End If


        If Total_Bottom_End / RebarLimit <= 1 Then
            member.bot1_E = Total_Bottom_End
            member.bot2_E = 0
            member.bot3_E = 0
        ElseIf Total_Bottom_Start / RebarLimit <= 2 Then
            member.bot1_E = RebarLimit
            member.bot2_E = Total_Bottom_End - RebarLimit
            member.bot3_E = 0
        ElseIf Total_Bottom_Start / RebarLimit <= 3 Then
            member.bot1_E = RebarLimit
            member.bot2_E = RebarLimit
            member.bot3_E = Total_Bottom_End - RebarLimit * 2
        Else
            'MsgBox("Over Maximum Allowable Value" & vbCrLf & "Member ID :  " & vbCrLf & member.OID)
            member.bot1_E = "Over"
            member.bot2_E = "Over"
            member.bot3_E = "Over"
            member.OverRebar = True
            'GoTo endFunc
        End If

        '================
        If Total_Top_Middle / RebarLimit <= 1 Then
            member.top1_M = Total_Top_Middle
            member.top2_M = 0
            member.top3_M = 0
        ElseIf Total_Top_Middle / RebarLimit <= 2 Then
            member.top1_M = RebarLimit
            member.top2_M = Total_Top_Middle - RebarLimit
            member.top3_M = 0
        ElseIf Total_Top_Middle / RebarLimit <= 3 Then
            member.top1_M = RebarLimit
            member.top2_M = RebarLimit
            member.top3_M = Total_Top_Middle - RebarLimit * 2
        Else
            'MsgBox("Over Maximum Allowable Value" & vbCrLf & "Member ID :  " & vbCrLf & member.OID)
            member.top1_M = "Over"
            member.top2_M = "Over"
            member.top3_M = "Over"
            member.OverRebar = True
            'GoTo endFunc
        End If


        If Total_Bottom_Middle / RebarLimit <= 1 Then
            member.bot1_M = Total_Bottom_Middle
            member.bot2_M = 0
            member.bot3_M = 0
        ElseIf Total_Bottom_Middle / RebarLimit <= 2 Then
            member.bot1_M = RebarLimit
            member.bot2_M = Total_Bottom_Middle - RebarLimit
            member.bot3_M = 0
        ElseIf Total_Bottom_Middle / RebarLimit <= 3 Then
            member.bot1_M = RebarLimit
            member.bot2_M = RebarLimit
            member.bot3_M = Total_Bottom_Middle - RebarLimit * 2
        Else
            'MsgBox("Over Maximum Allowable Value" & vbCrLf & "Member ID :  " & vbCrLf & member.OID)
            member.bot1_M = "Over"
            member.bot2_M = "Over"
            member.bot3_M = "Over"
            member.OverRebar = True
            'GoTo endFunc
        End If

        If member.OverRebar = True Then
            member.WarningMsg += "Rebar Number Over Maximum Allowable Value" & vbCrLf
            MsgBox("Rebar Number Over Maximum Allowable Value" & vbCrLf & vbCrLf & "Member ID  :  " & member.OID)
        End If

        '=======計算Side Bar 支數
        member.SideBar_Num = member.SideBar_Area / GetRebarArea(member.SideBar)

        If member.SideBar_Num Mod 2 <> 0 Then
            member.SideBar_Num += 1
        End If

        '=======角隅處需配置主筋
        If member.top1_S <= 1 Then member.top1_S = 2
        If member.top1_M <= 1 Then member.top1_M = 2
        If member.top1_E <= 1 Then member.top1_E = 2
        If member.bot1_S <= 1 Then member.bot1_S = 2
        If member.bot1_M <= 1 Then member.bot1_M = 2
        If member.bot1_E <= 1 Then member.bot1_E = 2

endFunc:
        Return member
    End Function

    '計算箍筋間距
    Public Function CalculateStirrup(ByRef Member As MemberProp, ByRef StirrupSize As String) As MemberProp

        Dim StirrupArea As Double = GetRebarArea(StirrupSize)

        Dim MaxMainBar As New List(Of Integer)
        MaxMainBar.Add(Member.top1_S)
        MaxMainBar.Add(Member.top1_M)
        MaxMainBar.Add(Member.top1_E)
        MaxMainBar.Add(Member.bot1_S)
        MaxMainBar.Add(Member.bot1_M)
        MaxMainBar.Add(Member.bot1_E)
        MaxMainBar.Sort()
        Dim MaxBarNumber As Integer = MaxMainBar.Max

        '=========箍筋10 cm 間距與支數檢核
        If Member.Stirrup_TTrnRebar_S * 10 > StirrupArea Then
            Member.WarningMsg += "Stirrup Size is too small for torsion bar (Start Section)" & vbCrLf
            GoTo Warning
        End If

        If (Member.Stirrup_VRebar_S + Member.Stirrup_TTrnRebar_S) * 10 / StirrupArea > Math.Max(CInt(Member.top1_S), CInt(Member.bot1_S)) Then
            Member.WarningMsg += "Stirrup Size is too small for torsion & Shear (Start Section)" & vbCrLf
        End If

        If (Member.Stirrup_VRebar_M + Member.Stirrup_TTrnRebar_M) * 10 / StirrupArea > Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.WarningMsg += "Stirrup Size is too small for torsion & Shear (Middle Section)" & vbCrLf
        End If

        If (Member.Stirrup_VRebar_E + Member.Stirrup_TTrnRebar_E) * 10 / StirrupArea > Math.Max(CInt(Member.top1_E), CInt(Member.bot1_E)) Then
            Member.WarningMsg += "Stirrup Size is too small for torsion & Shear (End Section)" & vbCrLf
        End If

        '==========計算扭剪箍筋支數 start 端 10/12/15 cm
        If Member.Stirrup_TTrnRebar_S * 10 < StirrupArea And Member.Stirrup_VRebar_S * 10 / StirrupArea <= MaxBarNumber Then
            Member.Stirrup_Num_S = Math.Ceiling((Member.Stirrup_TTrnRebar_S * 10 - StirrupArea + Member.Stirrup_VRebar_S * 10) / StirrupArea)
            Member.Stirrup_Spacing_S = 10
        End If
        If Member.Stirrup_TTrnRebar_S * 12 < StirrupArea And Member.Stirrup_VRebar_S * 12 / StirrupArea <= MaxBarNumber Then
            Member.Stirrup_Num_S = Math.Ceiling((Member.Stirrup_TTrnRebar_S * 12 - StirrupArea + Member.Stirrup_VRebar_S * 12) / StirrupArea)
            Member.Stirrup_Spacing_S = 12
        End If
        If Member.Stirrup_TTrnRebar_S * 15 < StirrupArea And Member.Stirrup_VRebar_S * 15 / StirrupArea <= MaxBarNumber Then
            Member.Stirrup_Num_S = Math.Ceiling((Member.Stirrup_TTrnRebar_S * 15 - StirrupArea + Member.Stirrup_VRebar_S * 15) / StirrupArea)
            Member.Stirrup_Spacing_S = 15
        End If

        If Member.Stirrup_Num_S > Math.Max(CInt(Member.top1_S), CInt(Member.bot1_S)) Then
            Member.WarningMsg += "Stirrup more than Main Bar Number(Start Section)" & vbCrLf
            'GoTo Warning
        End If

        '==========計算扭剪箍筋支數 end 端 10/12/15 cm
        If Member.Stirrup_TTrnRebar_E * 10 > StirrupArea Then
            Member.WarningMsg += "Stirrup Size is too small for torsion bar (End Section)" & vbCrLf
            GoTo Warning
        End If


        If Member.Stirrup_TTrnRebar_E * 10 < StirrupArea And Member.Stirrup_VRebar_E * 10 / StirrupArea <= MaxBarNumber Then
            Member.Stirrup_Num_E = Math.Ceiling((Member.Stirrup_TTrnRebar_E * 10 - StirrupArea + Member.Stirrup_VRebar_E * 10) / StirrupArea)
            Member.Stirrup_Spacing_E = 10
        End If
        If Member.Stirrup_TTrnRebar_E * 12 < StirrupArea And Member.Stirrup_VRebar_E * 12 / StirrupArea <= MaxBarNumber Then
            Member.Stirrup_Num_E = Math.Ceiling((Member.Stirrup_TTrnRebar_E * 12 - StirrupArea + Member.Stirrup_VRebar_E * 12) / StirrupArea)
            Member.Stirrup_Spacing_E = 12
        End If
        If Member.Stirrup_TTrnRebar_E * 15 < StirrupArea And Member.Stirrup_VRebar_E * 15 / StirrupArea <= MaxBarNumber Then
            Member.Stirrup_Num_E = Math.Ceiling((Member.Stirrup_TTrnRebar_E * 15 - StirrupArea + Member.Stirrup_VRebar_E * 15) / StirrupArea)
            Member.Stirrup_Spacing_E = 15
        End If


        If Member.Stirrup_Num_E > Math.Max(CInt(Member.top1_E), CInt(Member.bot1_E)) Then
            Member.WarningMsg += "Stirrup more than Main Bar Number(End Section)" & vbCrLf
            'GoTo Warning
        End If

        '==========計算扭剪箍筋支數 middle 端 10/12/15/18/20/22/25/30 cm
        If Member.Stirrup_TTrnRebar_M * 10 > StirrupArea Then
            Member.WarningMsg += "Stirrup Size is too small for torsion bar (Middle Section)" & vbCrLf
            'GoTo Warning
        End If

        Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 10 + Member.Stirrup_VRebar_M * 10) / StirrupArea)
        Member.Stirrup_Spacing_M = 10

        If Member.Stirrup_TTrnRebar_M * 10 < StirrupArea And Member.Stirrup_VRebar_M * 10 / StirrupArea <= Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 10 - StirrupArea + Member.Stirrup_VRebar_M * 10) / StirrupArea)
            Member.Stirrup_Spacing_M = 10
        End If
        If Member.Stirrup_TTrnRebar_M * 12 < StirrupArea And Member.Stirrup_VRebar_M * 12 / StirrupArea <= Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 12 - StirrupArea + Member.Stirrup_VRebar_M * 12) / StirrupArea)
            Member.Stirrup_Spacing_M = 12
        End If
        If Member.Stirrup_TTrnRebar_M * 15 < StirrupArea And Member.Stirrup_VRebar_M * 15 / StirrupArea <= Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 15 - StirrupArea + Member.Stirrup_VRebar_M * 15) / StirrupArea)
            Member.Stirrup_Spacing_M = 15
        End If
        If Member.Stirrup_TTrnRebar_M * 18 < StirrupArea And Member.Stirrup_VRebar_M * 18 / StirrupArea <= Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 18 - StirrupArea + Member.Stirrup_VRebar_M * 18) / StirrupArea)
            Member.Stirrup_Spacing_M = 18
        End If
        If Member.Stirrup_TTrnRebar_M * 20 < StirrupArea And Member.Stirrup_VRebar_M * 20 / StirrupArea <= Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 20 - StirrupArea + Member.Stirrup_VRebar_M * 20) / StirrupArea)
            Member.Stirrup_Spacing_M = 20
        End If
        If Member.Stirrup_TTrnRebar_M * 22 < StirrupArea And Member.Stirrup_VRebar_M * 22 / StirrupArea <= Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 22 - StirrupArea + Member.Stirrup_VRebar_M * 22) / StirrupArea)
            Member.Stirrup_Spacing_M = 22
        End If
        If Member.Stirrup_TTrnRebar_M * 25 < StirrupArea And Member.Stirrup_VRebar_M * 25 / StirrupArea <= Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 25 - StirrupArea + Member.Stirrup_VRebar_M * 25) / StirrupArea)
            Member.Stirrup_Spacing_M = 25
        End If
        If Member.Stirrup_TTrnRebar_M * 30 < StirrupArea And Member.Stirrup_VRebar_M * 30 / StirrupArea <= Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.Stirrup_Num_M = Math.Ceiling((Member.Stirrup_TTrnRebar_M * 30 - StirrupArea + Member.Stirrup_VRebar_M * 30) / StirrupArea)
            Member.Stirrup_Spacing_M = 30
        End If

        If Member.Stirrup_Num_M > Math.Max(CInt(Member.top1_M), CInt(Member.bot1_M)) Then
            Member.WarningMsg += "Stirrup more than Main Bar Number(Middle Section)" & vbCrLf
            'GoTo Warning
        End If

        '=======修正Stirrup支數(最少2支)
        If Member.Stirrup_Num_S <= 1 Then Member.Stirrup_Num_S = 2
        If Member.Stirrup_Num_M <= 1 Then Member.Stirrup_Num_M = 2
        If Member.Stirrup_Num_E <= 1 Then Member.Stirrup_Num_E = 2

Warning:
        Return CalculateStirrup
    End Function

    '取得Full Name 與 Piece Maek
    Public Function GetSizeMark(ByRef Member As MemberProp, ByRef sectionName As String) As MemberProp
        Dim divide As Integer

        divide = sectionName.IndexOf("-")

        If divide = -1 Then
            Member.FullName = sectionName
        Else
            Member.FullName = sectionName.Substring(divide + 1)
            Member.PieceMark = sectionName.Substring(0, divide)
        End If

        Return Member

    End Function


    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click

        Dim PiperackLoads_Point As New Piperack_Point_Loads

        PiperackLoads_Point.SapModel = SapModel
        PiperackLoads_Point.ISapPlugin = ISapPlugin
        PiperackLoads_Point.ShowDialog()
        ProgramStartCnt.inputName("SAP-Loading Input for PR(Point)")

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click

        Dim PiperackLoads_Distributed As New Piperack_Distributed_Loads

        PiperackLoads_Distributed.SapModel = SapModel
        PiperackLoads_Distributed.ISapPlugin = ISapPlugin
        PiperackLoads_Distributed.ShowDialog()
        ProgramStartCnt.inputName("SAP-Loading Input for PR(Dist.)")
    End Sub



    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        frmMenu = Nothing
        ISapPlugin.Finish(0)
        Me.Close()
    End Sub


    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click

        Dim SteelRatioList As New Steel_Ratio_List

        SteelRatioList.SapModel = SapModel
        SteelRatioList.ISapPlugin = ISapPlugin

        SteelRatioList.Show()
        ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Steel Ratio")
    End Sub



    Private Sub ComboBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComboBox1.MouseClick
        Dim selector As New OpenFileDialog
        If My.Computer.FileSystem.DirectoryExists("C:\Program Files\Computers and Structures\SAP2000 15\CTCIPlugin") Then
            selector.InitialDirectory = "C:\Program Files\Computers and Structures\SAP2000 15\CTCIPlugin\ssk"
        Else
            selector.InitialDirectory = "C:\Program Files (x86)\Computers and Structures\SAP2000 15\CTCIPlugin\ssk"
        End If


        selector.ShowDialog()

        Dim SkinName As String = selector.FileName

        ComboBox1.Text = SkinName
        SE.SkinFile = SkinName

        TextBox1.Text = SkinName.Replace("C:\Program Files\Computers and Structures\SAP2000 15\CTCIPlugin\ssk\", "")

    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click

        If SE.Active = True Then
            SkinEngine1.Active = False
        Else
            SE.Active = True
        End If
    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox3.Enter

    End Sub


    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
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

        If MsgBox("Unlock Model and Fix Memebr Start - End Joints", MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal, "Start - End Rule Check") = MsgBoxResult.Yes Then
            ret = SapModel.SetModelIsLocked(False)
        Else
            GoTo endSub
        End If

        ret = SapModel.SelectObj.All
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        Dim P1, P2 As String
        Dim X1, X2, Y1, Y2, Z1, Z2 As Double
        Dim P1P2NeedExchange As Boolean
        Dim ExCounter As Integer = 0


        For i = 0 To NumberItems - 1
            If ObjectType(i) = 2 Then
                ret = SapModel.FrameObj.GetPoints(ObjectName(i), P1, P2)

                ret = SapModel.PointObj.GetCoordCartesian(P1, X1, Y1, Z1)
                ret = SapModel.PointObj.GetCoordCartesian(P2, X2, Y2, Z2)


                'Zeor length member

                If Math.Abs(X1 - X2) < 0.00001 And Math.Abs(Y1 - Y2) < 0.00001 And Math.Abs(Z1 - Z2) < 0.00001 Then
                    MsgBox("Zero Length Member  :  " & ObjectName(i))
                End If

                'column
                If Math.Abs(X1 - X2) < 0.00001 And Math.Abs(Y1 - Y2) < 0.00001 And Math.Abs(Z1 - Z2) > 0.1 Then
                    If Z1 > Z2 Then P1P2NeedExchange = True
                End If

                'Beam (X Dir)
                If Math.Abs(Z1 - Z2) < 0.00001 And Math.Abs(X1 - X2) > 0.1 And Math.Abs(Y1 - Y2) < 0.00001 Then
                    If X1 > X2 Then P1P2NeedExchange = True
                End If

                'Baem  (Y Dir)
                If Math.Abs(Z1 - Z2) < 0.00001 And Math.Abs(Y1 - Y2) > 0.1 And Math.Abs(X1 - X2) < 0.00001 Then
                    If Y1 > Y2 Then P1P2NeedExchange = True
                End If

                'HB
                If Math.Abs(Z1 - Z2) < 0.00001 And Math.Abs(X1 - X2) > 0.1 And Math.Abs(Y1 - Y2) > 0.1 Then
                    If X1 > X2 Then P1P2NeedExchange = True
                End If

                'VB (XZ plane)
                If Math.Abs(Y1 - Y2) < 0.00001 And Math.Abs(X1 - X2) > 0.1 And Math.Abs(Z1 - Z2) > 0.1 Then
                    If X1 > X2 Then P1P2NeedExchange = True
                End If

                'VB(YZ plane)
                If Math.Abs(X1 - X2) < 0.00001 And Math.Abs(Y1 - Y2) > 0.1 And Math.Abs(Z1 - Z2) > 0.1 Then
                    If Y1 > Y2 Then P1P2NeedExchange = True
                End If

                'Other 
                If Math.Abs(X1 - X2) > 0.1 And Math.Abs(Y1 - Y2) > 0.1 And Math.Abs(Z1 - Z2) > 0.1 Then
                    If Z1 > Z2 Then P1P2NeedExchange = True
                End If

                If P1P2NeedExchange = True Then
                    ret = SapModel.EditFrame.ChangeConnectivity(ObjectName(i), P2, P1)
                    ExCounter += 1
                End If
                P1P2NeedExchange = False
            End If
        Next


        MsgBox(ExCounter & " member(s) fixed ", MsgBoxStyle.SystemModal, "Start - End Rule")

endSub:
        ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Start&End Rule Check")

    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click

        Dim Sidesway As New SideSway_Check

        Sidesway.SapModel = SapModel
        Sidesway.ISapPlugin = ISapPlugin

        Sidesway.Show()
        ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Sidesway Check")
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim frmWindArea As New WindArea

        frmWindArea.SapModel = SapModel
        frmWindArea.ISapPlugin = ISapPlugin

        frmWindArea.Show()
        ISapPlugin.Finish(0)
        ProgramStartCnt.inputName("SAP-Wind Load Area Cal.")

    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Dim frmPipingLoad As New Piping_Load_Import

        frmPipingLoad.SapModel = SapModel
        frmPipingLoad.ISapPlugin = ISapPlugin
        frmPipingLoad.Show()
        ISapPlugin.Finish(0)


    End Sub
    '2nd Order / 1St Order
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click

        'Dim frmCompareDrift As New Compare2ndAnd1stDraft
        'frmCompareDrift.sapmodel = SapModel



    End Sub

    Private Sub btnmodifyASD_CaseComb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnmodifyASD_CaseComb.Click

        Dim ret As Long
        Dim Count As Integer
        Dim NumberNames As Long
        Dim MyName() As String

        Dim MyLoadType() As String
        Dim MyLoadName() As String
        Dim MySF() As Double
        Dim NumberLoads As Long
        Dim LoadType() As String
        Dim LoadName() As String
        Dim SF() As Double

        Count = SapModel.LoadCases.Count
        ret = SapModel.LoadCases.GetNameList(NumberNames, MyName)

        For i = 0 To NumberNames - 1
            ret = SapModel.LoadCases.StaticNonlinear.GetLoads(MyName(i), NumberLoads, LoadType, LoadName, SF)
            If ret = 0 Then
                ret = SapModel.LoadCases.StaticNonlinear.SetCase("ASD" + MyName(i))
                ret = SapModel.RespCombo.Add(MyName(i) + "_ASD", 0)
                For j = 0 To SF.Count - 1
                    SF(j) = SF(j) * 1.6
                Next
                ret = SapModel.LoadCases.StaticNonlinear.SetLoads("ASD" + MyName(i), NumberLoads, LoadType, LoadName, SF)
                ret = SapModel.LoadCases.StaticNonlinear.SetGeometricNonlinearity("ASD" + MyName(i), 1)
                ret = SapModel.RespCombo.SetCaseList(MyName(i) + "_ASD", 0, "ASD" + MyName(i), 0.625)
            End If
        Next



    End Sub

    Private Sub fortest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fortest.Click
        Dim ret As Long
        Dim NumberNames As Long
        Dim GroupName() As String


        If MsgBox("Unlock Model and Continue", MsgBoxStyle.YesNo + MsgBoxStyle.SystemModal, "Auto Create Group") = MsgBoxResult.Yes Then
            ret = SapModel.SetModelIsLocked(False)
        Else
            GoTo endSub
        End If

        ret = SapModel.GroupDef.GetNameList(NumberNames, GroupName)

        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        For i = 0 To NumberItems - 1
            If ObjectType(i) = 5 Then
                ret = SapModel.AreaObj.SetLoadUniformToFrame("ALL", "DEAD", 0.01, 10, 2, False, "Global", eItemType.SelectedObjects)



            End If
        Next


endSub:

        ret = SapModel.SelectObj.ClearSelection
        ISapPlugin.Finish(0)
    End Sub



    Private Sub btnMemberEndForce_Click(sender As Object, e As EventArgs) Handles btnMemberEndForce.Click
        '還沒寫
        Dim ret As Long
        Dim NumberResults As Long
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


        Dim NumberCombs As Integer
        'Dim MyName2() As String
        'Dim Selected As Boolean
        Dim NumberItems As Integer
        Dim CombName() As String
        'Dim selectedCL As String

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
            MsgBox("Need Run Analysis First" & vbCrLf & "End Program", , "Deflection Check")
            GoTo ExitSub
        End If

        ret = SapModel.RespCombo.GetNameList(NumberCombs, CombName)


        ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput


        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        '=======20140715==============
        Dim NumberItems2 As Integer
        Dim ObjectType2() As Integer
        Dim ObjectName2() As String
        ret = SapModel.SelectObj.All
        ret = SapModel.SelectObj.GetSelected(NumberItems2, ObjectType2, ObjectName2)

        '清除與建立空的Group Dictionary
        GroupDict.Clear()
        For i = 0 To NumberItems2 - 1
            If ObjectType2(i) = 2 Then
                GroupDict.Add(ObjectName2(i), "")
            End If
        Next
        ret = SapModel.SelectObj.ClearSelection
        '=============================

        SapModel.SetPresentUnits(eUnits.Ton_m_C)

        'get member end force
        'Dim Obj() As String
        'Dim Elm() As String
        Dim PointElm() As String
        'Dim LoadCase() As String
        'Dim StepType() As String
        'Dim StepNum() As Double
        Dim F1() As Double
        Dim F2() As Double
        Dim F3() As Double
        Dim M1() As Double
        Dim M2() As Double
        Dim M3() As Double

        'set case and combo output selections
        ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")

        'get frame joint forces for line object "1"
        ret = SapModel.Results.FrameJointForce("1", 0, NumberResults, Obj, Elm, PointElm, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)



        'Dim SegmentCounter As Integer

        Dim SelectCombFrm As New SelectCombDialog

        For i = 0 To NumberCombs - 1
            SelectCombFrm.ListBox1.Items.Add(CombName(i))
        Next

        SelectCombFrm.ShowDialog()

        If SelectCombFlag = False Then GoTo ExitSub
        If OpenFileErrorFlag = True Then GoTo ExitSub


ExitSub:

    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click

        SaveFileDialog1.InitialDirectory = My.Computer.FileSystem.CurrentDirectory
        If SaveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
            GoTo endsub
        End If

        Dim PHEfile As String
        Dim NODfile As String

        PHEfile = SaveFileDialog1.FileName + ".PHE"
        NODfile = SaveFileDialog1.FileName + ".NOD"

        Dim ret As Long
        Dim NumberNames As Long
        Dim GroupName() As String

        Dim Name As String
        Dim NumberItems As Long
        Dim ObjectType() As Integer
        Dim ObjectName() As String

        FileOpen(10, PHEfile, OpenMode.Output)

        PrintLine(10, " ph num/line element num(s)")

        Dim groupstring As String

        ret = SapModel.GroupDef.GetNameList(NumberNames, GroupName)

        For i = 0 To NumberNames - 1
            If GroupName(i) <> "All" And IsNumeric(GroupName(i)) = True Then
                ret = SapModel.GroupDef.GetAssignments(GroupName(i), NumberItems, ObjectType, ObjectName)
                groupstring = GroupName(i).PadLeft(7, " ")
                For j = 0 To NumberItems - 1

                    If ObjectType(j) = 2 Then
                        groupstring = groupstring + ObjectName(j).PadLeft(7, " ")
                    End If
                Next
                PrintLine(10, groupstring)
                groupstring = ""
            End If
        Next

        FileClose(10)

        '=======================================================
        FileOpen(20, NODfile, OpenMode.Output)

        Dim NodeName() As String
        Dim x As Double, y As Double, z As Double
        Dim NodeString As String

        Dim Value() As Boolean
        Dim restraintStr As String = "  "
        Dim k() As Double

        SapModel.SetPresentUnits(eUnits.Ton_mm_C)

        ret = SapModel.PointObj.GetNameList(NumberNames, NodeName)
        PrintLine(20, NumberNames, NodeName(NumberNames - 1), "Unit :MM")

        For i = 0 To NumberNames - 1
            NodeString = NodeName(i).PadLeft(7, " ")
            ret = SapModel.PointObj.GetCoordCartesian(NodeName(i), x, y, z)
            ret = SapModel.PointObj.GetRestraint(NodeName(i), Value)

            If Value(0) = True Then restraintStr += "X"
            If Value(1) = True Then restraintStr += "Y"
            If Value(2) = True Then restraintStr += "Z"
            If Value(3) = True Then restraintStr += "RX"
            If Value(4) = True Then restraintStr += "RY"
            If Value(5) = True Then restraintStr += "RZ"

            ReDim k(5)
            ret = SapModel.PointObj.GetSpring(NodeName(i), k)

            If k(0) <> 0 Or k(1) <> 0 Or k(2) <> 0 Or k(3) <> 0 Or k(4) < 0 Or k(5) <> 0 Then
                restraintStr = "  Spring"
            End If

            NodeString = NodeString + x.ToString("F3").PadLeft(14, " ") + y.ToString("F3").PadLeft(12, " ") + z.ToString("F3").PadLeft(12, " ") + restraintStr

            PrintLine(20, NodeString)

            NodeString = ""
            restraintStr = "  "
        Next

        FileClose(20)





endsub:
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim frmKCalc As New K_Calc

        frmKCalc.SapModel = SapModel
        frmKCalc.ISapPlugin = ISapPlugin
        frmKCalc.Show()
        ISapPlugin.Finish(0)



    End Sub



    '示範如何抓與特定點相連接的objetc 或 element
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click

        'Dim SapObject As cOAPI
        'Dim SapModel As cSapModel




        Dim ret As Long
        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        Dim PointNumber() As Integer


        '建立Analysis model
        ret = SapModel.Analyze.CreateAnalysisModel


        'pointobj (抓分析前的model data)
        ret = SapModel.PointObj.GetConnectivity("14", NumberItems, ObjectType, ObjectName, PointNumber)

        'pointEle (抓分析model data)


        ret = SapModel.PointElm.GetConnectivity("14", NumberItems, ObjectType, ObjectName, PointNumber)



        ISapPlugin.Finish(0)


    End Sub

    Public RCBeamDesignRlt As New Dictionary(Of String, RCBeamData)
    Public RCColumnDesignRlt As New Dictionary(Of String, RCColumnData)


    '抽RC BEAM /Column DESIGN DATA
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        'Dim SapObject As cOAPI
        'Dim SapModel As cSapModel
        Dim MyUnits As eUnits

        Dim ret As Integer
        Dim Name As String
        Dim NumberItems As Integer
        Dim FrameName() As String
        MyUnits = SapModel.GetPresentUnits()

        'MsgBox(MyUnits.ToString)

        Dim ForceUnit, LengthUnit, temperatureUnit As String
        Dim Units() As String = MyUnits.ToString.Split("_")
        ForceUnit = Units(0)
        LengthUnit = Units(1)
        temperatureUnit = Units(2)


        Dim NumberCombs As Integer
        Dim CombName() As String
        Dim CaseName() As String
        Dim Status() As Integer
        Dim AnalyzedFlag As Boolean = False

        Dim ObjectType() As Integer
        Dim ObjectName() As String

        ret = SapModel.Analyze.GetCaseStatus(NumberItems, CaseName, Status)

        For i = 0 To NumberItems - 1
            If Status(i) = 4 Then
                AnalyzedFlag = True
            End If
        Next

        If AnalyzedFlag <> True Then
            MsgBox("Need To Run Analysis First" & vbCrLf & "End Program", , "Deflection Check")
            GoTo ExitSub
        End If

        '取得Comb List
        ret = SapModel.RespCombo.GetNameList(NumberCombs, CombName)
        '清除現在Result選擇的Case&Comb
        ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput

        'ret = SapModel.SelectObj.All
        '選擇Group : All Frame
        ret = SapModel.FrameObj.SetSelected("All", True, eItemType.Group)

        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        'check 是否是桿件、是否是RC，是否有設計結果 ，是梁設計或柱設計
        For i = 0 To NumberItems - 1

            If CheckFrameisRC(ObjectName(i)) <> True Then
                GoTo NextMember
            End If

            'MsgBox("過第1關 :斷面名稱帶RC")

            If getRC_ResultsBeam(ObjectName(i)) <> True Then
                'MsgBox("Frame :" & ObjectName(i) & "has no RC Beam design result,please check!")
                'GoTo nextmember
            End If

            If getRC_ResultsColumn(ObjectName(i)) <> True Then
                'MsgBox("Frame :" & ObjectName(i) & "has no RC column design result,please check!")
                GoTo nextmember
            End If

            'MsgBox("過第2關")


            Try
                CBToolKit.SerializeHelper.Serializer.SerializeToJsonFile(RCBeamDesignRlt(ObjectName(i)), "d:\tesT.JSON")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try




            'CBToolKit.SerializeHelper.Serializer.SerializeToJsonFile(RCBeamDesignRlt(ObjectName(i)), "D:\test.json")



NextMember:
        Next

        '輸出Beam Data 



ExitSub:
    End Sub

    Public Function CheckFrameisRC(ByRef FrameLabel As String) As Boolean
        Dim ret As Long
        Dim NumberNames As Long
        Dim GroupName() As String

        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String

        Dim PropName As String
        Dim SAuto As String


        ret = SapModel.FrameObj.GetSection(FrameLabel, PropName, SAuto)  'Section Name

        '==========以名稱判斷是否為RC
        If PropName Is Nothing Then
            CheckFrameisRC = False
            Return CheckFrameisRC
        End If

        If PropName.Contains("RC") = False Then
            CheckFrameisRC = False
        Else
            CheckFrameisRC = True
        End If
        Return CheckFrameisRC

    End Function

    Public Structure RCBeamData
        Dim ret As Long
        Dim Name As String
        Dim NumberItems As Long
        Dim FrameName() As String
        Dim Location() As Double
        Dim TopCombo() As String
        Dim TopArea() As Double
        Dim BotCombo() As String
        Dim BotArea() As Double
        Dim VmajorCombo() As String
        Dim VmajorArea() As Double
        Dim TLCombo() As String
        Dim TLArea() As Double
        Dim TTCombo() As String
        Dim TTArea() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String
    End Structure

    Public Function getRC_ResultsBeam(ByRef FrameLabel As String) As Boolean
        Dim ret As Long
        Dim Name As String
        Dim NumberItems As Long
        Dim FrameName() As String
        Dim Location() As Double
        Dim TopCombo() As String
        Dim TopArea() As Double
        Dim BotCombo() As String
        Dim BotArea() As Double
        Dim VmajorCombo() As String
        Dim VmajorArea() As Double
        Dim TLCombo() As String
        Dim TLArea() As Double
        Dim TTCombo() As String
        Dim TTArea() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String

        Dim BeamDataCol As New RCBeamData

        'get Beam summary result data
        ret = SapModel.DesignConcrete.GetSummaryResultsBeam(FrameLabel, NumberItems, FrameName, Location, TopCombo, TopArea, BotCombo, BotArea, VmajorCombo, VmajorArea, TLCombo, TLArea, TTCombo, TTArea, ErrorSummary, WarningSummary)


        If ret = 1 Or FrameName Is Nothing Then
            getRC_ResultsBeam = False
        Else
            getRC_ResultsBeam = True

            BeamDataCol.Name = FrameLabel
            BeamDataCol.NumberItems = NumberItems
            BeamDataCol.FrameName = FrameName
            BeamDataCol.Location = Location
            BeamDataCol.TopCombo = TopCombo
            BeamDataCol.TopArea = TopArea
            BeamDataCol.BotCombo = BotCombo
            BeamDataCol.BotArea = BotArea
            BeamDataCol.VmajorCombo = VmajorCombo
            BeamDataCol.VmajorArea = VmajorArea
            BeamDataCol.TLCombo = TLCombo
            BeamDataCol.TLArea = TLArea
            BeamDataCol.TTCombo = TTCombo
            BeamDataCol.TTArea = TTArea
            BeamDataCol.ErrorSummary = ErrorSummary
            BeamDataCol.WarningSummary = WarningSummary

            RCBeamDesignRlt.Add(FrameLabel, BeamDataCol)

        End If
        Return getRC_ResultsBeam
    End Function

    Public Structure RCColumnData
        Dim ret As Long
        Dim Name As String
        Dim NumberItems As Long
        Dim FrameName() As String
        Dim MyOption() As Integer
        Dim Location() As Double
        Dim PMMCombo() As String
        Dim PMMArea() As Double
        Dim PMMRatio() As Double
        Dim VmajorCombo() As String
        Dim AVmajor() As Double
        Dim VminorCombo() As String
        Dim AVminor() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String
    End Structure

    Public Structure FrameForceQuake
        Dim CombName As String
        Dim WithQuake As Boolean
        Dim CaseList() As String
        Dim PatternType() As Object
        Dim ScaleFactor() As Double
        Dim P() As Double
        Dim V2() As Double
        Dim V3() As Double
        Dim T() As Double
        Dim M2() As Double
        Dim M3() As Double
    End Structure


    Public Structure RCDesignResult
        '共通
        Dim LengthUnit, ForceUnit As String
        Dim framelabel As String
        Dim group() As String
        Dim piecemark As String
        Dim sectionProfile As String
        Dim width, depth As Double
        Dim length, NetLength, StartOffsetLength, EndOffsetLength As Double


        'Load Case Factor
        'Dim CombwithQuake() As String
        'Dim LoadCaseList() As String
        'Dim LoadCaseSF() As String
        '載種組合中包含地震力，帶係數的Dead load + Live load 
        Dim FrameForceQuakes() As FrameForceQuake

        Dim isBeamorColumn As String

        'if beam

        Dim BEAM_CoverTop, BEAM_CoverBot As Double
        Dim BEAM_TopLeftArea As Double
        Dim BEAM_TopRightArea As Double
        Dim BEAM_BotLeftArea As Double
        Dim BEAM_BotRightArea As Double

        'Dim FrameName() As String
        Dim Location() As Double
        Dim TopCombo() As String
        Dim TopArea() As Double
        Dim BotCombo() As String
        Dim BotArea() As Double
        Dim VmajorCombo() As String
        Dim VmajorArea() As Double
        Dim TLCombo() As String
        Dim TLArea() As Double
        Dim TTCombo() As String
        Dim TTArea() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String




        'if column

        Dim COLUMN_Pattern As Long
        Dim COLUMN_ConfineType As Long
        Dim COLUMN_ClearCoverforConfinementBars As Double
        Dim COLUMN_NumberCBars As Long
        Dim COLUMN_NumberR3Bars As Long
        Dim COLUMN_NumberR2Bars As Long
        Dim COLUMN_RebarSize As String
        Dim COLUMN_TieSize As String
        Dim COLUMN_TieSpacingLongit As Double
        Dim COLUMN_Number2DirTieBars As Long
        Dim COLUMN_Number3DirTieBars As Long
        Dim COLUMN_ToBeDesigned As Boolean
        'Dim ret As Long
        'Dim Name As String
        'Dim NumberItems As Long
        'Dim MyOption() As Integer
        Dim PMMCombo() As String
        Dim PMMArea() As Double
        Dim PMMRatio() As Double
        Dim AVmajor() As Double
        Dim VminorCombo() As String
        Dim AVminor() As Double

        'Beam & Column Design force
        Dim NumberResults As Long
        Dim ComboName() As String
        Dim Station() As Double
        Dim P() As Double
        Dim V2() As Double
        Dim V3() As Double
        Dim T() As Double
        Dim M2() As Double
        Dim M3() As Double


    End Structure



    Public Function getRC_ResultsColumn(ByRef FrameLabel As String) As Boolean
        Dim ret As Long
        Dim Name As String
        Dim NumberItems As Long
        Dim FrameName() As String
        Dim MyOption() As Integer
        Dim Location() As Double
        Dim PMMCombo() As String
        Dim PMMArea() As Double
        Dim PMMRatio() As Double
        Dim VmajorCombo() As String
        Dim AVmajor() As Double
        Dim VminorCombo() As String
        Dim AVminor() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String

        Dim ColumnDataCol As New RCColumnData

        'get Column summary result data
        ret = SapModel.DesignConcrete.GetSummaryResultsColumn(FrameLabel, NumberItems, FrameName, MyOption, Location, PMMCombo, PMMArea, PMMRatio, VmajorCombo, AVmajor, VminorCombo, AVminor, ErrorSummary, WarningSummary)

        If ret = 1 Or FrameName Is Nothing Then
            getRC_ResultsColumn = False
        Else
            getRC_ResultsColumn = True

            ColumnDataCol.Name = FrameLabel
            ColumnDataCol.NumberItems = NumberItems
            ColumnDataCol.FrameName = FrameName
            ColumnDataCol.MyOption = MyOption
            ColumnDataCol.Location = Location
            ColumnDataCol.PMMCombo = PMMCombo
            ColumnDataCol.PMMArea = PMMArea
            ColumnDataCol.PMMRatio = PMMRatio
            ColumnDataCol.VmajorCombo = VmajorCombo
            ColumnDataCol.AVmajor = AVmajor
            ColumnDataCol.VminorCombo = VminorCombo
            ColumnDataCol.AVminor = AVminor
            ColumnDataCol.ErrorSummary = ErrorSummary
            ColumnDataCol.WarningSummary = WarningSummary

            RCColumnDesignRlt.Add(FrameLabel, ColumnDataCol)

        End If

        Return getRC_ResultsColumn





    End Function


    Public Function getFrameDesignType(ByRef FrameLabel As String) As String



    End Function

    Public Function getFrameMaterial(ByRef FrameLabel As String) As String



        Dim ret As Long
        Dim PropName As String
        Dim SAuto As String
        ret = SapModel.FrameObj.GetSection(FrameLabel, PropName, SAuto)

        Dim FileName As String
        Dim MatProp As String
        Dim t3 As Double
        Dim t2 As Double
        Dim Color As Long
        Dim Notes As String
        Dim GUID As String

        'get frame section property data
        ret = SapModel.PropFrame.GetRectangle(PropName, FileName, MatProp, t3, t2, Color, Notes, GUID)

        Dim MatType As eMatType
        'MatType = eMatType.NoDesign

        'Dim Color_ As Long
        'Dim Notes_ As String
        'Dim GUID_ As String


        ret = SapModel.PropMaterial.GetMaterial(MatProp, MatType, Color, Notes, GUID)

        Dim ShapeName() As String
        Dim MyType() As Integer
        Dim DesignType As Long
        Dim NumberItems As Long

        'get section designer section property data
        ret = SapModel.PropFrame.GetSDSection(PropName, MatProp, NumberItems, ShapeName, MyType, DesignType, Color, Notes, GUID)

        Dim NumberResults As Long
        Dim FrameName() As String
        Dim ComboName() As String
        Dim Station() As Double
        Dim P() As Double
        Dim V2() As Double
        Dim V3() As Double
        Dim T() As Double
        Dim M2() As Double
        Dim M3() As Double

        'get beam design forces
        ret = SapModel.DesignResults.DesignForces.BeamDesignForces(FrameLabel, NumberResults, FrameName, ComboName, Station, P, V2, V3, T, M2, M3, eItemType.Objects)

        'get column design forces
        ret = SapModel.DesignResults.DesignForces.ColumnDesignForces(FrameLabel, NumberResults, FrameName, ComboName, Station, P, V2, V3, T, M2, M3, eItemType.Objects)

        'get design section
        ret = SapModel.DesignConcrete.GetDesignSection(FrameLabel, PropName)

        Dim MyName() As String


        'get combos selected for concrete strength design
        ret = SapModel.DesignConcrete.GetComboStrength(NumberItems, MyName)


        Dim CType1() As eCNameType
        Dim CName() As String
        Dim SF() As Double

        'get all cases and combos included in combo COMB104
        ret = SapModel.RespCombo.GetCaseList("COMB104", NumberItems, CType1, CName, SF)




        Dim Obj() As String
        Dim ObjSta() As Double
        Dim Elm() As String
        Dim ElmSta() As Double
        Dim LoadCase() As String
        Dim StepType() As String
        Dim StepNum() As Double

        'get frame forces for line object "FrameLabel"


        ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
        ret = SapModel.Results.Setup.SetCaseSelectedForOutput("eqp")

        ret = SapModel.Results.FrameForce(FrameLabel, eItemTypeElm.ObjectElm, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)





        MsgBox(SapModel.DesignConcrete.GetResultsAvailable)

        Dim Name As String
        Dim MyOption() As Integer
        Dim Location() As Double
        Dim PMMCombo() As String
        Dim PMMArea() As Double
        Dim PMMRatio() As Double
        Dim VmajorCombo() As String
        Dim AVmajor() As Double
        Dim VminorCombo() As String
        Dim AVminor() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String


        ret = SapModel.DesignConcrete.GetSummaryResultsColumn(FrameLabel, NumberItems, FrameName, MyOption, Location, PMMCombo, PMMArea, PMMRatio, VmajorCombo, AVmajor, VminorCombo, AVminor, ErrorSummary, WarningSummary)


        Dim TopCombo() As String
        Dim TopArea() As Double
        Dim BotCombo() As String
        Dim BotArea() As Double

        Dim VmajorArea() As Double
        Dim TLCombo() As String
        Dim TLArea() As Double
        Dim TTCombo() As String
        Dim TTArea() As Double

        'get summary result data
        ret = SapModel.DesignConcrete.GetSummaryResultsBeam(FrameLabel, NumberItems, FrameName, Location, TopCombo, TopArea, BotCombo, BotArea, VmajorCombo, VmajorArea, TLCombo, TLArea, TTCombo, TTArea, ErrorSummary, WarningSummary)


        Dim n1 As Integer
        Dim n2 As Integer
        'Dim MyName() As String

        'verify frame objects successfully designed
        'n1 : design NG n2 : not designed RC members
        ret = SapModel.DesignConcrete.VerifyPassed(NumberItems, n1, n2, MyName)

        Dim myType1 As Integer

        '1 = Steel 2 = Concrete
        ret = SapModel.FrameObj.GetDesignProcedure(FrameLabel, myType1)

        Dim NumberGroups As Long
        Dim Groups() As String


        'get frame object groups
        ret = SapModel.FrameObj.GetGroupAssign(FrameLabel, NumberGroups, Groups)

        Dim AutoOffset As Boolean
        Dim Length1 As Double = 0
        Dim Length2 As Double = 0
        Dim rz As Double = 0


        'get offsets for line element
        ret = SapModel.LineElm.GetEndLengthOffset(FrameLabel, Length1, Length2, rz)

        'get offsets
        ret = SapModel.FrameObj.GetEndLengthOffset(FrameLabel, AutoOffset, Length1, Length2, rz)

        Dim nelm As Long
        'Dim Elm() As String
        Dim RDI() As Double
        Dim RDJ() As Double


        'get object information for a line element
        ret = SapModel.FrameObj.GetElm(FrameLabel, nelm, Elm, RDI, RDJ)

        Dim Point1 As String
        Dim Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double
        Dim memberLength As Double

        'get names of points
        ret = SapModel.FrameObj.GetPoints(FrameLabel, Point1, Point2)
        'get point coordinates
        ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
        ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)
        memberLength = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5


        Dim MatPropLong As String
        Dim MatPropConfine As String
        Dim CoverTop As Double
        Dim CoverBot As Double
        Dim TopLeftArea As Double
        Dim TopRightArea As Double
        Dim BotLeftArea As Double
        Dim BotRightArea As Double


        'get beam rebar data (用來取柱子還是會回傳0，無法用這判別BEAM or COLUMN)
        ret = SapModel.PropFrame.GetRebarBeam(PropName, MatPropLong, MatPropConfine, CoverTop, CoverBot, TopLeftArea, TopRightArea, BotLeftArea, BotRightArea)

        Dim Pattern As Long
        Dim ConfineType As Long
        Dim Cover As Double
        Dim NumberCBars As Long
        Dim NumberR3Bars As Long
        Dim NumberR2Bars As Long
        Dim RebarSize As String
        Dim TieSize As String
        Dim TieSpacingLongit As Double
        Dim Number2DirTieBars As Long
        Dim Number3DirTieBars As Long
        Dim ToBeDesigned As Boolean

        'get column rebar data
        ret = SapModel.PropFrame.GetRebarColumn(PropName, MatPropLong, MatPropConfine, Pattern, ConfineType, Cover, NumberCBars, NumberR3Bars, NumberR2Bars, RebarSize, TieSize, TieSpacingLongit, Number2DirTieBars, Number3DirTieBars, ToBeDesigned)




        Return MatType.ToString

    End Function



    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        MsgBox(getFrameMaterial("00777777"))
        'Dim jsonStr As String = My.Computer.FileSystem.ReadAllText("D:\RCDATAEXTRACT.json")
        'Dim jsonDict As Dictionary(Of String, RCDesignResult)
        'Try
        '    'Dim newDict As Dictionary(Of String, RCDesignResult) = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Dictionary(Of String, RCDesignResult))(jsonStr)
        '    jsonDict = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Dictionary(Of String, RCDesignResult))(jsonStr)
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try



    End Sub

    '以桿件幾何關係判定是COLUMN或BEAM
    Public Function IdentifyType(ByRef FrameLabel As String) As String


        Dim ret As Long
        Dim Point1, Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        'get names of points
        ret = SapModel.FrameObj.GetPoints(FrameLabel, Point1, Point2)
        'get point coordinates
        ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
        ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)

        If Math.Abs(X1 - X2) < 0.1 And Math.Abs(Y1 - Y2) < 0.1 And Math.Abs(Z1 - Z2) > 0.1 Then
            IdentifyType = "COLUMN"
        Else
            IdentifyType = "BEAM"
        End If

        Return IdentifyType

    End Function

    Public Function isConcrete(ByRef FrameLabel As String) As Boolean
        Dim ret As Long
        Dim MyType As Long
        'get design procedure
        ret = SapModel.FrameObj.GetDesignProcedure(FrameLabel, MyType)
        If MyType = 2 Then
            isConcrete = True
        Else
            isConcrete = False
        End If
        Return isConcrete
    End Function

    'Public RCDesignResultCol As New Dictionary(Of String, RCDesignResult)

    Private Sub ExtractRCData_Click(sender As Object, e As EventArgs) Handles ExtractRCData.Click

        Dim RCDesignResultCol As New Dictionary(Of String, RCDesignResult)


        Dim RCmemberDgnRlt As New RCDesignResult



        Dim ret As Long
        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String

        Dim MyType As Long

        Dim MyUnits As eUnits

        MyUnits = SapModel.GetPresentUnits()

        Dim ForceUnit, LengthUnit, temperatureUnit As String
        Dim Units() As String = MyUnits.ToString.Split("_")

        ForceUnit = Units(0)
        LengthUnit = Units(1)
        temperatureUnit = Units(2)


        Dim NumberGroups As Long
        Dim Groups() As String

        Dim MatPropLong As String
        Dim MatPropConfine As String
        Dim CoverTop As Double
        Dim CoverBot As Double
        Dim TopLeftArea As Double
        Dim TopRightArea As Double
        Dim BotLeftArea As Double
        Dim BotRightArea As Double


        Dim Name As String
        Dim FrameName() As String
        Dim MyOption() As Integer
        Dim Location() As Double
        Dim PMMCombo() As String
        Dim PMMArea() As Double
        Dim PMMRatio() As Double
        Dim VmajorCombo() As String
        Dim AVmajor() As Double
        Dim VminorCombo() As String
        Dim AVminor() As Double
        Dim ErrorSummary() As String
        Dim WarningSummary() As String

        Dim TopCombo() As String
        Dim TopArea() As Double
        Dim BotCombo() As String
        Dim BotArea() As Double

        Dim VmajorArea() As Double
        Dim TLCombo() As String
        Dim TLArea() As Double
        Dim TTCombo() As String
        Dim TTArea() As Double

        Dim NumberResults As Long
        Dim ComboName() As String
        Dim Station() As Double
        Dim P() As Double
        Dim V2() As Double
        Dim V3() As Double
        Dim T() As Double
        Dim M2() As Double
        Dim M3() As Double

        Dim FileName As String
        Dim MatProp As String
        Dim t3 As Double
        Dim t2 As Double
        Dim Color As Long
        Dim Notes As String
        Dim GUID As String

        Dim AutoOffset As Boolean
        Dim Length1 As Double = 0
        Dim Length2 As Double = 0
        Dim rz As Double = 0

        Dim Point1 As String
        Dim Point2 As String
        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double
        Dim memberLength As Double

        Dim Pattern As Long
        Dim ConfineType As Long
        Dim Cover As Double
        Dim NumberCBars As Long
        Dim NumberR3Bars As Long
        Dim NumberR2Bars As Long
        Dim RebarSize As String
        Dim TieSize As String
        Dim TieSpacingLongit As Double
        Dim Number2DirTieBars As Long
        Dim Number3DirTieBars As Long
        Dim ToBeDesigned As Boolean

        Dim SectionProfile As String

        Dim MyName() As String

        Dim CType1() As eCNameType
        Dim CName() As String
        Dim SF() As Double

        Dim NumberLoads As Long
        Dim LoadType() As String
        Dim LoadName() As String


        Dim MyType1 As eLoadPatternType
        Dim PatternType As eLoadPatternType
        Dim PatternCnt As Integer = 0
        Dim ForceCounter As Integer = 0

        Dim Obj() As String
        Dim ObjSta() As Double
        Dim Elm() As String
        Dim ElmSta() As Double
        Dim LoadCase() As String
        Dim StepType() As String
        Dim StepNum() As Double

        If SapModel.DesignConcrete.GetResultsAvailable = False Then
            MsgBox("You Need To Perform Design Concrete First!")
            Exit Sub
        End If

        'get combos selected for concrete strength design
        ret = SapModel.DesignConcrete.GetComboStrength(NumberItems, MyName)
        '先預定陣列大小等於comb總數
        ReDim RCmemberDgnRlt.FrameForceQuakes(NumberItems - 1)
        For Each CombName In MyName
            '取得裡面的load case name
            ret = SapModel.RespCombo.GetCaseList(CombName, NumberItems, CType1, CName, SF)

            For i = 0 To NumberItems - 1
                '如果是Non-linear case ，則取出裡面的pattern
                If SapModel.LoadCases.StaticNonlinear.GetLoads(CName(i), NumberLoads, LoadType, LoadName, SF) = 0 And CType1(i) = eCNameType.LoadCase Then
                    For Each LN In LoadName
                        ret = SapModel.LoadPatterns.GetLoadType(LN, MyType1)
                        '檢查每個pattern，如果包含Quake則標記並輸出
                        If MyType1 = eLoadPatternType.Quake Then
                            'ReDim RCmemberDgnRlt.FrameForceQuakes(NumberLoads - 1)
                            RCmemberDgnRlt.FrameForceQuakes(ForceCounter).CombName = CombName
                            RCmemberDgnRlt.FrameForceQuakes(ForceCounter).WithQuake = True
                            RCmemberDgnRlt.FrameForceQuakes(ForceCounter).CaseList = LoadName
                            RCmemberDgnRlt.FrameForceQuakes(ForceCounter).ScaleFactor = SF

                            '指定PatternType 陣列大小
                            ReDim Preserve RCmemberDgnRlt.FrameForceQuakes(ForceCounter).PatternType(LoadName.Length - 1)

                            For Each LoadpatternType In LoadName
                                ret = SapModel.LoadPatterns.GetLoadType(LoadpatternType, PatternType)
                                RCmemberDgnRlt.FrameForceQuakes(ForceCounter).PatternType(PatternCnt) = PatternType.ToString
                                PatternCnt += 1
                            Next
                            '重置counter
                            PatternCnt = 0


                            'RCmemberDgnRlt.FrameForceQuakes(ForceCounter).Station = Station
                            'ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
                            'ret = SapModel.Results.Setup.SetCaseSelectedForOutput(LoadName(0))
                            ForceCounter += 1
                        End If
                    Next
                End If
            Next
        Next
        '程序結束，重新redim 陣列大小
        ReDim Preserve RCmemberDgnRlt.FrameForceQuakes(ForceCounter - 1)



        'get selected objects
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
        If NumberItems = 0 Then
            ret = SapModel.SelectObj.All
            ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
        End If

        For i = 0 To NumberItems - 1
            '2 = Frame object
            If ObjectType(i) = 2 And isConcrete(ObjectName(i)) Then

                '單位
                RCmemberDgnRlt.ForceUnit = Units(0)
                RCmemberDgnRlt.LengthUnit = Units(1)

                'get frame object groups (梁柱共通)
                ret = SapModel.FrameObj.GetGroupAssign(ObjectName(i), NumberGroups, Groups)
                RCmemberDgnRlt.group = Groups

                'get design section
                ret = SapModel.DesignConcrete.GetDesignSection(ObjectName(i), SectionProfile)

                RCmemberDgnRlt.framelabel = ObjectName(i)
                RCmemberDgnRlt.sectionProfile = SectionProfile

                'get frame section property data (Width / Depth)
                ret = SapModel.PropFrame.GetRectangle(SectionProfile, FileName, MatProp, t3, t2, Color, Notes, GUID)
                RCmemberDgnRlt.width = t2
                RCmemberDgnRlt.depth = t3

                'get offsets
                ret = SapModel.FrameObj.GetEndLengthOffset(ObjectName(i), AutoOffset, Length1, Length2, rz)
                RCmemberDgnRlt.StartOffsetLength = Length1
                RCmemberDgnRlt.EndOffsetLength = Length2

                'get names of points
                ret = SapModel.FrameObj.GetPoints(ObjectName(i), Point1, Point2)
                'get point coordinates
                ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)
                ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)
                memberLength = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5
                RCmemberDgnRlt.length = Format(memberLength, "0.000")
                RCmemberDgnRlt.NetLength = Format(memberLength - Length1 - Length2, "0.000")

                '計算QUAKE相關載重



                For j = 0 To RCmemberDgnRlt.FrameForceQuakes.Length - 1
                    For k = 0 To RCmemberDgnRlt.FrameForceQuakes(j).PatternType.Length - 1

                        If RCmemberDgnRlt.FrameForceQuakes(j).PatternType(k) = "Dead" Or RCmemberDgnRlt.FrameForceQuakes(j).PatternType(k) = "Live" Then
                            ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
                            ret = SapModel.Results.Setup.SetCaseSelectedForOutput(RCmemberDgnRlt.FrameForceQuakes(j).CaseList(k))

                            ret = SapModel.Results.FrameForce(ObjectName(i), eItemTypeElm.ObjectElm, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)

                            ReDim Preserve RCmemberDgnRlt.FrameForceQuakes(j).P(NumberResults - 1)
                            ReDim Preserve RCmemberDgnRlt.FrameForceQuakes(j).V2(NumberResults - 1)
                            ReDim Preserve RCmemberDgnRlt.FrameForceQuakes(j).V3(NumberResults - 1)
                            ReDim Preserve RCmemberDgnRlt.FrameForceQuakes(j).T(NumberResults - 1)
                            ReDim Preserve RCmemberDgnRlt.FrameForceQuakes(j).M2(NumberResults - 1)
                            ReDim Preserve RCmemberDgnRlt.FrameForceQuakes(j).M3(NumberResults - 1)

                            RCmemberDgnRlt.FrameForceQuakes(j).P = RCmemberDgnRlt.FrameForceQuakes(j).P.Zip(P, Function(x, y) x + y * RCmemberDgnRlt.FrameForceQuakes(j).ScaleFactor(k)).ToArray
                            RCmemberDgnRlt.FrameForceQuakes(j).V2 = RCmemberDgnRlt.FrameForceQuakes(j).V2.Zip(V2, Function(x, y) x + y * RCmemberDgnRlt.FrameForceQuakes(j).ScaleFactor(k)).ToArray
                            RCmemberDgnRlt.FrameForceQuakes(j).V3 = RCmemberDgnRlt.FrameForceQuakes(j).V3.Zip(V3, Function(x, y) x + y * RCmemberDgnRlt.FrameForceQuakes(j).ScaleFactor(k)).ToArray
                            RCmemberDgnRlt.FrameForceQuakes(j).T = RCmemberDgnRlt.FrameForceQuakes(j).T.Zip(T, Function(x, y) x + y * RCmemberDgnRlt.FrameForceQuakes(j).ScaleFactor(k)).ToArray
                            RCmemberDgnRlt.FrameForceQuakes(j).M2 = RCmemberDgnRlt.FrameForceQuakes(j).M2.Zip(M2, Function(x, y) x + y * RCmemberDgnRlt.FrameForceQuakes(j).ScaleFactor(k)).ToArray
                            RCmemberDgnRlt.FrameForceQuakes(j).M3 = RCmemberDgnRlt.FrameForceQuakes(j).M3.Zip(M3, Function(x, y) x + y * RCmemberDgnRlt.FrameForceQuakes(j).ScaleFactor(k)).ToArray



                        End If
                    Next

                Next

                'ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
                'ret = SapModel.Results.Setup.SetCaseSelectedForOutput("LL")
                'ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DL")
                'ret = SapModel.Results.FrameForce(ObjectName(i), eItemTypeElm.ObjectElm, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)



                'RCmemberDgnRlt.FrameForceQuakes(0).P = P
                'RCmemberDgnRlt.FrameForceQuakes(0).V2 = V2
                'RCmemberDgnRlt.FrameForceQuakes(0).V3 = V3
                'RCmemberDgnRlt.FrameForceQuakes(0).T = T
                'RCmemberDgnRlt.FrameForceQuakes(0).M2 = M2
                'RCmemberDgnRlt.FrameForceQuakes(0).M3 = M3

                'RCmemberDgnRlt.FrameForceQuakes(0).P = Array.ConvertAll(P, Function(x) x * 1000)

                'RCmemberDgnRlt.FrameForceQuakes(0).P = Array.ConvertAll(Of Double, Double)(P, Convert.ToDouble(P))

                'RCmemberDgnRlt.FrameForceQuakes(0).P = Convert.ToDouble(P) * 100

                'RCmemberDgnRlt.FrameForceQuakes(0).P=






                If IdentifyType(ObjectName(i)) = "COLUMN" Then

                    'get column rebar data
                    ret = SapModel.PropFrame.GetRebarColumn(SectionProfile, MatPropLong, MatPropConfine, Pattern, ConfineType, Cover, NumberCBars, NumberR3Bars, NumberR2Bars, RebarSize, TieSize, TieSpacingLongit, Number2DirTieBars, Number3DirTieBars, ToBeDesigned)
                    RCmemberDgnRlt.COLUMN_Pattern = Pattern
                    RCmemberDgnRlt.COLUMN_ConfineType = ConfineType
                    RCmemberDgnRlt.COLUMN_ClearCoverforConfinementBars = Cover
                    RCmemberDgnRlt.COLUMN_NumberCBars = NumberCBars
                    RCmemberDgnRlt.COLUMN_NumberR3Bars = NumberR3Bars
                    RCmemberDgnRlt.COLUMN_NumberR2Bars = NumberR2Bars
                    RCmemberDgnRlt.COLUMN_RebarSize = RebarSize
                    RCmemberDgnRlt.COLUMN_TieSize = TieSize
                    RCmemberDgnRlt.COLUMN_TieSpacingLongit = TieSpacingLongit
                    RCmemberDgnRlt.COLUMN_Number2DirTieBars = Number2DirTieBars
                    RCmemberDgnRlt.COLUMN_Number3DirTieBars = Number3DirTieBars
                    RCmemberDgnRlt.COLUMN_ToBeDesigned = ToBeDesigned


                    'get column design forces
                    ret = SapModel.DesignResults.DesignForces.ColumnDesignForces(ObjectName(i), NumberResults, FrameName, ComboName, Station, P, V2, V3, T, M2, M3, eItemType.Objects)
                    RCmemberDgnRlt.NumberResults = NumberResults
                    RCmemberDgnRlt.ComboName = ComboName
                    RCmemberDgnRlt.Station = Station
                    RCmemberDgnRlt.P = P
                    RCmemberDgnRlt.V2 = V2
                    RCmemberDgnRlt.V3 = V3
                    RCmemberDgnRlt.T = T
                    RCmemberDgnRlt.M2 = M2
                    RCmemberDgnRlt.M3 = M3

                    'get column design result
                    ret = SapModel.DesignConcrete.GetSummaryResultsColumn(ObjectName(i), NumberItems, FrameName, MyOption, Location, PMMCombo, PMMArea, PMMRatio, VmajorCombo, AVmajor, VminorCombo, AVminor, ErrorSummary, WarningSummary)
                    RCmemberDgnRlt.Location = Location
                    RCmemberDgnRlt.PMMCombo = PMMCombo
                    RCmemberDgnRlt.PMMArea = PMMArea
                    RCmemberDgnRlt.PMMRatio = PMMRatio
                    RCmemberDgnRlt.VmajorCombo = VmajorCombo
                    RCmemberDgnRlt.AVmajor = AVmajor
                    RCmemberDgnRlt.VminorCombo = VminorCombo
                    RCmemberDgnRlt.AVminor = AVminor
                    RCmemberDgnRlt.ErrorSummary = ErrorSummary
                    RCmemberDgnRlt.WarningSummary = WarningSummary

                    RCmemberDgnRlt.isBeamorColumn = "COLUMN"
                ElseIf IdentifyType(ObjectName(i)) = "BEAM" Then

                    'get beam rebar data
                    ret = SapModel.PropFrame.GetRebarBeam(SectionProfile, MatPropLong, MatPropConfine, CoverTop, CoverBot, TopLeftArea, TopRightArea, BotLeftArea, BotRightArea)
                    RCmemberDgnRlt.BEAM_CoverTop = CoverTop
                    RCmemberDgnRlt.BEAM_CoverBot = CoverBot
                    RCmemberDgnRlt.BEAM_TopLeftArea = TopLeftArea
                    RCmemberDgnRlt.BEAM_TopRightArea = TopRightArea
                    RCmemberDgnRlt.BEAM_BotLeftArea = BotLeftArea
                    RCmemberDgnRlt.BEAM_BotRightArea = BotRightArea

                    'get beam design forces
                    ret = SapModel.DesignResults.DesignForces.BeamDesignForces(ObjectName(i), NumberResults, FrameName, ComboName, Station, P, V2, V3, T, M2, M3, eItemType.Objects)
                    RCmemberDgnRlt.NumberResults = NumberResults
                    RCmemberDgnRlt.ComboName = ComboName
                    RCmemberDgnRlt.Station = Station
                    RCmemberDgnRlt.P = P
                    RCmemberDgnRlt.V2 = V2
                    RCmemberDgnRlt.V3 = V3
                    RCmemberDgnRlt.T = T
                    RCmemberDgnRlt.M2 = M2
                    RCmemberDgnRlt.M3 = M3
                    'get summary result data
                    ret = SapModel.DesignConcrete.GetSummaryResultsBeam(ObjectName(i), NumberItems, FrameName, Location, TopCombo, TopArea, BotCombo, BotArea, VmajorCombo, VmajorArea, TLCombo, TLArea, TTCombo, TTArea, ErrorSummary, WarningSummary)
                    RCmemberDgnRlt.Location = Location
                    RCmemberDgnRlt.TopCombo = TopCombo
                    RCmemberDgnRlt.TopArea = TopArea
                    RCmemberDgnRlt.BotCombo = BotCombo
                    RCmemberDgnRlt.BotArea = BotArea
                    RCmemberDgnRlt.VmajorCombo = VmajorCombo
                    RCmemberDgnRlt.VmajorArea = VmajorArea
                    RCmemberDgnRlt.TLCombo = TLCombo
                    RCmemberDgnRlt.TLArea = TLArea
                    RCmemberDgnRlt.TTCombo = TTCombo
                    RCmemberDgnRlt.TTArea = TTArea
                    RCmemberDgnRlt.ErrorSummary = ErrorSummary
                    RCmemberDgnRlt.WarningSummary = WarningSummary

                    RCmemberDgnRlt.isBeamorColumn = "BEAM"
                End If

                RCDesignResultCol.Add(ObjectName(i), RCmemberDgnRlt)
            End If
        Next

        Dim jsonStr As String = ""




        '================Save Json File=====================
        Dim SaveFileDia As New SaveFileDialog
        Dim JsonFileName As String

        SaveFileDia.FileName = "RC Design Data"
        SaveFileDia.DefaultExt = "json"
        SaveFileDia.Filter = "RC Design Data(*.json) |*.json"


        If SaveFileDia.ShowDialog() = DialogResult.OK Then
            JsonFileName = SaveFileDia.FileName
        Else
            GoTo Notselectfile
        End If
        '================End Save Json File=================

        Try
            'CBToolKit.SerializeHelper.Serializer.SerializeToJsonFile(RCDesignResultCol, "D:\RCDATAtoJSON1.txt")
            jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(RCDesignResultCol)
            'My.Computer.FileSystem.WriteAllText("D:\RCDATAEXTRACT.json", jsonStr, False)
            'Process.Start("D:\RCDATAEXTRACT.json")
            My.Computer.FileSystem.WriteAllText(JsonFileName, jsonStr, False)
            Process.Start(JsonFileName)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
Notselectfile:


    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click

        Dim ret As Long
        Dim RebarName As String

        RebarName = "A615Gr60"

        SapModel.SetPresentUnits(eUnits.Ton_mm_C)
        ret = SapModel.PropFrame.SetRebarBeam("RC400X600", RebarName, RebarName, 35, 30, 4100, 4200, 4300, 4400)




    End Sub
End Class

'v1.12 2013/05/09 RC Beam design result export
'v1.15 2013/06/20 deflection 條件修改 (重大)
'v1.16 2013/06/28 deflection 新增對象過濾條件(功能改善)
'v1.17 2013/07/18 deflection 新增檢核實際相對變位數值(功能增加)
'v1.20 2013/09/10 RC Deisgn  功能新增(未完成)
'v1.25 2014/02/13 Regular Point Loads 功能新增
'v1.26 2014/02/27 Distributed Loads 功能新增
'v1.30 2014/03/10 Steel Ratio List 功能新增
'v1.31 2014/03/14 Start End Rule Check 功能新增
'v1.33 2014/03/24 Steel Ratio List 介面修改
'v1.35 2014/05/16 Sidesway check 功能新增
'v1.40 2014/06/10 Steel Ratio List Group讀取效能改善
'v1.41 2014/06/17 Wind Area 功能新增
'v1.44 2014/11/28 Delete group 功能修改
'v1.46 2015/06/08 point load & distributed load 可儲存上次輸入之數值
'v1.5  2015/09/10 新增各功能說明文檔連結 
'v1.6  2016/02/18 RC PIECE MARK分類程式修改
'v1.7  2016/04/21 RC Beam分類程式修改elevation編號規則 (EL:0 判別)
'v1.8  2016/09/19 Sidesway 程式修改(SAP2000 節點座標精度問題
'v1.9  2016/09/26 modify nonlinear case/comb for ASD
'...
'v20.1 2019/5/28  RC 分析設計資料抽取&序列化/反序列化
