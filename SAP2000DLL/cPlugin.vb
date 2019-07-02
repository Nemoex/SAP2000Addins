'Imports Sap2000
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports Sunisoft.IrisSkin
Imports SAP2000v20


Module Module1
    Public xlApp As Microsoft.Office.Interop.Excel.Application
    Public xlBook As Workbook
    Public xlSheet As Worksheet
    Public xlRange As Range
    Public frmMenu As Menu
    Public SE As New SkinEngine

End Module




Namespace CreateGroup
    Public Class cPlugin
        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

            'If System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower <> "ctci.com.tw" Then
            '    MsgBox("CTCI Only")
            '    GoTo endSub
            'End If

            Dim ret As Long
            Dim NumberNames As Long
            Dim GroupName() As String

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
            MsgBox("Group Complete", , "Create Group")
endSub:

            ret = SapModel.SelectObj.ClearSelection
            ISapPlugin.Finish(0)
        End Sub

    End Class


End Namespace

Namespace GroupByList
    Public Class cPlugin
        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)
            'Dim NumberItems As Integer = 0
            'Dim ObjectTypes As Array = Nothing
            'Dim ObjectNames As Array = Nothing
            'SapModel.SelectObj.GetSelected(NumberItems, ObjectTypes, ObjectNames)
            MsgBox("for test only")
            Dim openfile As New OpenFileDialog
            openfile.ShowDialog()

            If openfile.FileName = "" Then
                ISapPlugin.Finish(0)
                Exit Sub
            End If


            Dim elementlist As String = openfile.FileName

            FileOpen(10, elementlist, OpenMode.Input)



            Dim elementNo As String
            elementNo = InputBox("Input element number")

            If elementNo = "" Then
                FileClose(10)
                ISapPlugin.Finish(0)
                Exit Sub
            End If



            Dim GroupName As String
            GroupName = InputBox("Input Group Name")

            If GroupName = "" Then
                FileClose(10)
                ISapPlugin.Finish(0)
                Exit Sub
            End If



            Dim List(CInt(elementNo) - 1) As Integer

            Dim i As Integer = 0
            Dim listdata As String
            Do Until EOF(10)
                listdata = LineInput(10)
                If listdata = "" Then Exit Do
                List(i) = listdata
                i += 1
            Loop


            FileClose(10)


            Dim ret As Long
            'Dim label As String
            ret = SapModel.GroupDef.SetGroup(GroupName)
            For j = 1 To i
                ret = SapModel.FrameObj.SetGroupAssign(List(j - 1), GroupName)

            Next

            MsgBox("Group :   " & GroupName & "   Created " & vbCrLf & "Total  : " & i & "  elements")



            ISapPlugin.Finish(0)
        End Sub

    End Class


End Namespace

Namespace deleteGroup
    Public Class cPlugin
        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

            'If System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower <> "ctci.com.tw" Then
            '    MsgBox("CTCI Only")
            '    GoTo endSub
            'End If

            Dim ret As Long
            Dim NumberNames As Long
            Dim GroupName() As String

            ret = SapModel.GroupDef.GetNameList(NumberNames, GroupName)

            For i = 0 To NumberNames - 1
                If GroupName(i) <> "All" Then
                    ret = SapModel.GroupDef.Delete(GroupName(i))
                End If
            Next

            MsgBox("Delete All Group Complete!")
endSub:
            ISapPlugin.Finish(0)
        End Sub

    End Class

End Namespace

Namespace ReadTable
    Public Class cPlugin
        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)
            'Dim NumberItems As Integer = 0
            'Dim ObjectTypes As Array = Nothing
            'Dim ObjectNames As Array = Nothing
            'SapModel.SelectObj.GetSelected(NumberItems, ObjectTypes, ObjectNames)

            'If System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower <> "ctci.com.tw" Then
            '    MsgBox("CTCI Only")
            '    GoTo endSub
            'End If

            Dim SapObject As cOAPI
            'Dim SapModel As cSapModel
            Dim ret As Long
            'Dim Name As String
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



            Dim MyUnits As eUnits
            'Dim TempUnits As eUnits



            MyUnits = SapModel.GetPresentUnits

            'switch length to cm
            SapModel.SetPresentUnits(eUnits.kgf_cm_C)



            'ret = SapModel.PropFrame.SetRectangle("RC400X600A", "4000Psi", 60, 40)

            'ret = SapModel.PropFrame.SetRebarBeam("RC400X600A", "A615Gr60", "A615Gr40", 3.5, 3.6, 4.1, 4.2, 4.3, 4.4)

            'ret = SapModel.PropFrame.SetRectangle("R2", "4000Psi", 60, 40)

            'ret = SapModel.PropFrame.SetRebarColumn("R2", "A615Gr60", "A615Gr40", 1, 1, 5.8, 0, 6, 7, "#10", "#5", 17, 4, 5, False)

            ret = SapModel.Analyze.RunAnalysis


            'NumberItems = 1
            ret = SapModel.DesignConcrete.StartDesign

            'get summary result data
            ret = SapModel.DesignConcrete.GetSummaryResultsBeam("AGroup", NumberItems, FrameName, Location, TopCombo, TopArea, BotCombo, BotArea, VmajorCombo, VmajorArea, TLCombo, TLArea, TTCombo, TTArea, ErrorSummary, WarningSummary, eItemType.Group)

            'Dim NumberName As Integer
            'Dim MyName() As String

            Dim ObjectType() As Integer
            Dim ObjectName() As String


            'ret = SapModel.GroupDef.GetNameList(NumberName, MyName)

            ret = SapModel.GroupDef.GetAssignments("AGroup", NumberItems, ObjectType, ObjectName)

            Dim LCJSRatioMajor() As String
            Dim JSRatioMajor() As Double
            Dim LCJSRatioMinor() As String
            Dim JSRatioMinor() As Double
            Dim LCBCCRatioMajor() As String
            Dim BCCRatioMajor() As Double
            Dim LCBCCRatioMinor() As String
            Dim BCCRatioMinor() As Double
            'Dim ErrorSummary() As String
            'Dim WarningSummary() As String

            ret = SapModel.DesignConcrete.GetSummaryResultsJoint("2", NumberItems, FrameName, LCJSRatioMajor, JSRatioMajor, LCJSRatioMinor, JSRatioMinor, LCBCCRatioMajor, BCCRatioMajor, LCBCCRatioMinor, BCCRatioMinor, ErrorSummary, WarningSummary)



            'MsgBox(FrameName(0).ToString)

            'close Sap2000
            'SapObject.ApplicationExit(False)

            'switch unit back
            SapModel.SetPresentUnits(MyUnits)



            SapModel = Nothing
            SapObject = Nothing


endSub:
            ISapPlugin.Finish(0)
        End Sub

    End Class


End Namespace

Namespace AutoCreatePool



    Public Class cPlugin



        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)
            'If CheckDomain() = False Then
            '    MsgBox("This Plugin is for CTCI Only")
            '    Exit Sub
            'End If

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

endSub:     ISapPlugin.Finish(0)
        End Sub

    End Class


End Namespace

Namespace SplitColumn

    Public Class cPlugin
        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)
            Dim ret As Long
            Dim ParallelTo(5) As Boolean
            Dim Num As Long
            Dim NewName() As String
            Dim NumberItems As Long
            Dim ObjectType() As Integer
            Dim ObjectName() As String
            Dim memberCounter As Integer

            'ret = SapModel.SelectObj.All
            ParallelTo(2) = True
            ret = SapModel.SelectObj.LinesParallelToCoordAxis(ParallelTo)
            ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

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
            ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)




            MsgBox(memberCounter & "  Members Divided" & _
                   vbCrLf & ObjectName.Count & "  New Members Created", MsgBoxStyle.Information, "Complete")

            ISapPlugin.Finish(0)


        End Sub
    End Class

End Namespace

Namespace CheckBeamDepth

    Public Class cPlugin
        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

            'If CheckDomain() = False Then
            '    MsgBox("This Plugin is for CTCI Only")
            '    Exit Sub
            'End If

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
                If PropName.Contains("H") Then
                    ret = SapModel.PropFrame.GetISection(PropName, FileName, MatProp, t3, t2, tf, tw, t2b, tfb, Color, Notes, GUID)
                ElseIf PropName.Contains("C") Then
                    ret = SapModel.PropFrame.GetChannel(PropName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
                ElseIf PropName.Contains("T") Then
                    ret = SapModel.PropFrame.GetTee(PropName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
                End If

                If t3 = 0 Then
                    UnknowFrame(i) = ObjectName(i)

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
                    CantRecognizrList = CantRecognizrList + UnknowFrame(i) + ","
                    UnknowCounter += 1
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
                MsgBox("All Members are OK!", , "Check Beam Depth")
                PrintLine(60, "")
                PrintLine(60, "All Members Depth Are Greater Then Length/30 ")
                PrintLine(60, "")
                PrintLine(60, "          =================")
                PrintLine(60, "          =      O K      =")
                PrintLine(60, "          =================")

            Else
                MsgBox("Frame depth less then length/30  :" & vbCrLf & Result & _
                       vbCrLf & vbCrLf & ProblemFrameCounter & "   Frames selected", MsgBoxStyle.Critical, "Check Beam Depth")
                PrintLine(60, "")
                PrintLine(60, ProblemFrameCounter & " Frame(s) Depth Are Less Then Length/30 ")
                PrintLine(60, "")
                PrintLine(60, "Frame ID : ")
                PrintLine(60, Result)
                PrintLine(60, "")
                PrintLine(60, "          =================")
                PrintLine(60, "          =    Caution    =")
                PrintLine(60, "          =================")
            End If

            MsgBox(resultFile, , "Check Result File")
            FileClose(60)

            ISapPlugin.Finish(0)

        End Sub
    End Class

End Namespace

Namespace DeflectionCheck


    Public Class cPlugin

        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)



            'If CheckDomain() = False Then
            '    MsgBox("This Plugin is for CTCI Only")
            '    GoTo exitsub
            'End If

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
                Else
                    MsgBox("Please select frame first then re-start command", , "Deflection Check")
                    GoTo ExitSub
                End If
            End If

            Dim SelectCombFrm As New SelectCombDialog

            For i = 0 To NumberCombs - 1
                SelectCombFrm.ListBox1.Items.Add(CombName(i))
            Next

            SelectCombFrm.ShowDialog()

            If SelectCombFlag = False Then GoTo ExitSub
            If OpenFileErrorFlag = True Then GoTo exitsub

            Dim SelectedFrameCount As Integer
            ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)
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
            PrintLine(15, "   Phy No /CTRL Comb/ Deflection / ******** / OK-NG / Critical Node /Node List")
            'PrintLine(15, "                        mm          mm")
            PrintLine(15, "")
            PrintLine(15, "======================================================================")

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

                    '==========2014/01/03 避免Group name 沒變更的問題
                    If Elm Is Nothing Then
                        ret = SapModel.Results.JointDispl(ObjectName(m), eItemTypeElm.ObjectElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
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

                If VoidPointCounter <> NumSegs And UseVNodeFlag = False Then
                    VoidNumCompare = " (" + (NumSegs - VoidPointCounter).ToString + " Points Overlap" + ")"
                Else
                    VoidNumCompare = ""

                End If


                If DefChkCriteria = True Then
                    PrintLine(15, "    " + ObjectName(m), Microsoft.VisualBasic.TAB(13), CombList(MaxCombFlag), Microsoft.VisualBasic.TAB(25), "L/" + CStr(CInt(MaxResult)) + "   ", Microsoft.VisualBasic.TAB(48), OKFlag, Microsoft.VisualBasic.TAB(60), Elm(MaxNodeFlag), Microsoft.VisualBasic.TAB(68), "  |  " + NodeList + VoidNumCompare)
                Else
                    PrintLine(15, "    " + ObjectName(m), Microsoft.VisualBasic.TAB(13), CombList(MaxCombFlag), Microsoft.VisualBasic.TAB(25), " " + CStr(Format(CDbl(MaxResult), "0.00")) + "   ", Microsoft.VisualBasic.TAB(48), OKFlag, Microsoft.VisualBasic.TAB(60), Elm(MaxNodeFlag), Microsoft.VisualBasic.TAB(68), "  |  " + NodeList + VoidNumCompare)
                End If


NextItem:   Next


            MsgBox("Complete" & vbCrLf & ReportFileName, , "Deflection Check")

ExitSub:
            FileClose(15)
            ISapPlugin.Finish(0)


        End Sub

        Public Function fraction(ByRef StartPoint As Point3D, ByRef EndPoint As Point3D, ByVal MidPoint As Point3D, ByVal SDef As Double, ByVal EDef As Double, ByVal MDef As Double) As Double

            Dim length1, length2 As Double
            Dim linearDef As Double
            Dim RelDef As Double

            length1 = ((StartPoint.X - MidPoint.X) ^ 2 + (StartPoint.Y - MidPoint.Y) ^ 2 + (StartPoint.Z - MidPoint.Z) ^ 2) ^ 0.5
            length2 = ((EndPoint.X - MidPoint.X) ^ 2 + (EndPoint.Y - MidPoint.Y) ^ 2 + (EndPoint.Z - MidPoint.Z) ^ 2) ^ 0.5

            linearDef = EDef + (SDef - EDef) / (length1 + length2) * length2
            RelDef = MDef - linearDef
            fraction = Math.Abs(1 / (RelDef / (length1 + length2)))


        End Function

        Structure Point3D
            Dim X As Double
            Dim Y As Double
            Dim Z As Double
        End Structure


    End Class
End Namespace

Namespace OutputReaction
    Public Class cPlugin

        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

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
                MsgBox("Need Run Analysis First" & vbCrLf & "End Program", , "Reaction")
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
        End Sub

        Private Function AlignS(ByRef Num As String, Optional ByVal Dec As Integer = 3, Optional ByRef Space As Integer = 10) As String

            Dim Output As String
            Num = CDbl(Num)
            Output = FormatNumber(Num, Dec).PadLeft(Space)

            AlignS = Output.ToString
            Return AlignS

        End Function

    End Class
End Namespace

Namespace ExportMDT
    Public Class cPlugin
        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

            'Dim SaveFileDia As New SaveFileDialog

            'SaveFileDia.FileName = ""
            'SaveFileDia.DefaultExt = "MDT"
            'SaveFileDia.Filter = "Model Data(*.MDT) |*.MDT"

            'If SaveFileDia.ShowDialog() = DialogResult.OK Then
            '    FileOpen(100, SaveFileDia.FileName, OpenMode.Output)
            'Else
            '    GoTo exitsub
            'End If

            Dim ret As Long

            SapModel.SetPresentUnits(eUnits.Ton_mm_C)
            ret = SapModel.SelectObj.All
            Dim NumberItems As Integer
            Dim ObjectType() As Integer
            Dim ObjectName() As String
            ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

            Dim StartP, EndP As String
            Dim StartX, StartY, StartZ As Double
            Dim EndX, EndY, EndZ As Double
            Dim CP As Integer
            Dim Mirror As Boolean
            Dim StiffTransform As Boolean
            Dim Offset1() As Double
            Dim Offset2() As Double
            Dim CSys As String

            Dim ii() As Boolean
            Dim jj() As Boolean
            Dim StartValue() As Double
            Dim EndValue() As Double

            Dim PropName As String
            Dim SAuto As String

            Dim Ang As Double
            Dim Advanced As Boolean

            Dim NumberGroups As Long
            Dim GroupName() As String

            'Dim GroupContains As Integer
            Dim GNumberItems As Integer
            Dim GObjectType() As Integer
            Dim GObjectName() As String
            Dim RecordMax As Integer

            For k = 0 To NumberItems - 1
                If ObjectName(k) > RecordMax Then
                    RecordMax = ObjectName(k)
                End If
            Next

            Dim MemberGroupList(RecordMax) As String

            'Get ALL group name and record im Myname list
            ret = SapModel.GroupDef.GetNameList(NumberGroups, GroupName)
            'get group assignments
            For i = 0 To NumberGroups - 1
                If GroupName(i).ToLower = "all" Then GoTo NextGroup
                If IsNumeric(GroupName(i)) = False Then GoTo nextgroup

                ret = SapModel.GroupDef.GetAssignments(GroupName(i), GNumberItems, GObjectType, GObjectName)
                For j = 0 To GNumberItems - 1
                    If GObjectType(j) = 2 Then
                        MemberGroupList(GObjectName(j)) = GroupName(i).ToString
                    End If
                Next
NextGroup:
            Next


            Dim MemberGroupName As String

            Dim FileName As String
            Dim MatProp As String
            Dim t3 As Double
            Dim t2 As Double
            Dim tf As Double
            Dim tw As Double
            Dim t2b As Double
            Dim tfb As Double
            Dim dis As Double
            Dim Color As Long
            Dim Notes As String
            Dim GUID As String
            Dim UsageType As String

            Dim PropType As eFramePropType
            Dim MatType As eMatType

            Dim Material As String

            For i = 0 To NumberItems - 1
                If ObjectType(i) = 2 Then
                    'Get Section Name
                    ret = SapModel.FrameObj.GetSection(ObjectName(i), PropName, SAuto)
                    'Get member start&end point ID
                    ret = SapModel.FrameObj.GetPoints(ObjectName(i), StartP, EndP)
                    'Get start&end point coordinate XYZ value 
                    ret = SapModel.PointObj.GetCoordCartesian(StartP, StartX, StartY, StartZ)
                    ret = SapModel.PointObj.GetCoordCartesian(EndP, EndX, EndY, EndZ)
                    'Get member CP / Mirror / stiff / Offsetij / CS
                    ret = SapModel.FrameObj.GetInsertionPoint(ObjectName(i), CP, Mirror, StiffTransform, Offset1, Offset2, CSys)
                    'Get member end release condition
                    ret = SapModel.FrameObj.GetReleases(ObjectName(i), ii, jj, StartValue, EndValue)
                    'Get Local Axes (rotation angle)
                    ret = SapModel.FrameObj.GetLocalAxes(ObjectName(i), Ang, Advanced)
                    'Get member Group Name
                    MemberGroupName = MemberGroupList(ObjectName(i))
                    'Get Usage Type
                    UsageType = DetermineType(StartX, StartY, StartZ, EndX, EndY, EndZ)

                    'Get Section Type
                    ret = SapModel.PropFrame.GetTypeOAPI(PropName, PropType)

                    'Get section Material & section dimension
                    Select Case PropType
                        Case eFramePropType.I
                            ret = SapModel.PropFrame.GetISection(PropName, FileName, MatProp, t3, t2, tf, tw, t2b, tfb, Color, Notes, GUID)
                        Case eFramePropType.Channel
                            ret = SapModel.PropFrame.GetChannel(PropName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
                        Case eFramePropType.T
                            ret = SapModel.PropFrame.GetTee(PropName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
                        Case eFramePropType.Angle
                            ret = SapModel.PropFrame.GetAngle(PropName, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
                        Case eFramePropType.DblAngle
                            ret = SapModel.PropFrame.GetDblAngle(PropName, FileName, MatProp, t3, t2, tf, tw, dis, Color, Notes, GUID)
                        Case eFramePropType.Box
                            'ret - SapModel.PropFrame .get
                            ret = 1
                            MatProp = " "
                        Case eFramePropType.Pipe
                            ret = SapModel.PropFrame.GetPipe(PropName, FileName, MatProp, t3, tw, Color, Notes, GUID)
                        Case eFramePropType.Rectangular
                            ret = SapModel.PropFrame.GetRectangle(PropName, FileName, MatProp, t3, t2, Color, Notes, GUID)
                        Case eFramePropType.DbChannel
                            ret = SapModel.PropFrame.GetDblChannel(PropName, FileName, MatProp, t3, t2, tf, tw, dis, Color, Notes, GUID)
                        Case Else
                            MsgBox("Not in Case 1~8")
                            MatProp = " "
                    End Select


                    ret = SapModel.PropMaterial.GetMaterial(MatProp, MatType, Color, Notes, GUID)
                    Select Case MatType
                        Case eMatType.Steel
                            Material = "STEEL"
                        Case eMatType.Concrete
                            Material = "CONCRETE"
                        Case eMatType.Rebar
                            Material = "REBAR"
                        Case eMatType.Aluminum
                            Material = "ALUMINUM"
                        Case Else
                            Material = "OTHER"
                    End Select



                    FileOpen(100, "V:\54833\KK.txt", OpenMode.Output)
                    PrintLine(100, MemberGroupName & " :" & StartP & "," & EndP & " " & UsageType & "--------" & "ST")




                End If

            Next


exitsub:
            ret = SapModel.SelectObj.ClearSelection
            ISapPlugin.Finish(0)
        End Sub

        Private Function DetermineType(ByRef Sx As Double, ByRef Sy As Double, ByRef Sz As Double, ByRef Ex As Double, ByRef Ey As Double, ByRef Ez As Double) As String

            '輸入起點/終點座標 透過規則定出該桿件類別
            '分為四類 C / B / HB / VB
            If Sx = Ex And Sy = Ey And Math.Abs(Ez - Sz) > 0 Then
                Return "C"
            End If
            If Sz = Ez Then
                If (Sx <> Ex And Sy - Ey = 0) Or (Sy <> Ey And Sx - Ex = 0) Then
                    Return "B"
                ElseIf Sx <> Ex And Sy <> Ey Then
                    Return "HB"
                End If
            End If
            If (Sz <> Ez And Sx <> Ex) Or (Sz <> Ez And Sy <> Ey) Then
                Return "VB"
            End If
            Return "OT"
        End Function


    End Class

End Namespace


Namespace GetSummaryResultBeam
    Public Class cPlugin

        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

            Dim ret As Long
            'Dim Name As String
            Dim NumberItems As Integer
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


            Dim Name As String

            Dim MyOption() As Integer

            Dim PMMCombo() As String
            Dim PMMArea() As Double
            Dim PMMRatio() As Double

            Dim AVmajor() As Double
            Dim VminorCombo() As String
            Dim AVminor() As Double




            Dim NumberNames As Long
            Dim MyName() As String



            'ret = SapModel.PropFrame.GetNameList(NumberNames, MyName)
            ret = SapModel.SelectObj.All

            Dim NumberItemsA As Integer
            Dim ObjectType() As Integer
            Dim ObjectName() As String
            ret = SapModel.SelectObj.GetSelected(NumberItemsA, ObjectType, ObjectName)
            ret = SapModel.SelectObj.All

            'ret =SapModel.SelectObj .

            'ret = SapModel.DesignConcrete.GetSummaryResultsColumn("2", NumberItems, FrameName, MyOption, Location, PMMCombo, PMMArea, PMMRatio, VmajorCombo, AVmajor, VminorCombo, AVminor, ErrorSummary, WarningSummary)

            Dim retr As Long

            'For i = 0 To NumberItemsA - 1
            '    If ObjectType(i) = 2 Then
            '        retr = SapModel.DesignConcrete.GetSummaryResultsColumn(ObjectName(i), NumberItems, FrameName, MyOption, Location, PMMCombo, PMMArea, PMMRatio, VmajorCombo, AVmajor, VminorCombo, AVminor, ErrorSummary, WarningSummary)

            '        retr = SapModel.DesignConcrete.GetSummaryResultsBeam(ObjectName(i), NumberItems, FrameName, Location, TopCombo, TopArea, BotCombo, BotArea, VmajorCombo, VmajorArea, TLCombo, TLArea, TTCombo, TTArea, ErrorSummary, WarningSummary)
            '    End If
            'Next

            Dim RebarName As String
            ret = SapModel.PropFrame.SetRebarBeam("1GX1", "A615Gr60", "A615Gr40", 3.5, 3, 4.1, 4.2, 4.3, 4.4)






            ret = SapModel.SelectObj.ClearSelection
            ISapPlugin.Finish(0)
        End Sub

    End Class
End Namespace


Namespace LoadXlsData
    Public Class cPlugin

        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

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

            Dim Str1, Str2, Str3 As String
            Dim TableName As String

            TableName = xlSheet.Cells(1, 1).value.ToString

            Dim LineCounter As Integer = 4

            Dim SectCounter As Integer = 1
            Dim SectNumber As Integer

            Dim StartSect As Integer
            Dim EndSect As Integer

            Dim CoverTop, CoverBot As Double
            Dim FTopAreaStart, FBotAreaStart As Double
            Dim FTopAreaEnd, FBotAreaEnd As Double


            Dim SectName As String



            If TableName.Contains("Concrete Design") And TableName.Contains("Beam Summary Data") Then
                Do Until xlSheet.Cells(LineCounter, 1).Value = Nothing
                    LineCounter += 1
                Loop
            End If

            Dim ret As Long
            'Dim RebarName As String

            'CoverTop = 5.0
            'CoverBot = 5.0
            '=======================================================
            SapModel.SetPresentUnits(eUnits.kgf_cm_C)


            Dim InputForm As New Form()

            Dim Label1 = New System.Windows.Forms.Label
            Dim ComboBox1 = New System.Windows.Forms.ComboBox
            Dim Label2 = New System.Windows.Forms.Label
            Dim ComboBox2 = New System.Windows.Forms.ComboBox
            Dim Label3 = New System.Windows.Forms.Label
            Dim Label4 = New System.Windows.Forms.Label
            Dim TextBox1 = New System.Windows.Forms.TextBox
            Dim TextBox2 = New System.Windows.Forms.TextBox
            Dim Label5 = New System.Windows.Forms.Label
            Dim Label6 = New System.Windows.Forms.Label
            Dim Button1 = New System.Windows.Forms.Button
            Dim Label7 = New System.Windows.Forms.Label
            Dim Combobox3 = New System.Windows.Forms.ComboBox

            InputForm.SuspendLayout()
            '
            'Label1
            '
            Label1.AutoSize = True
            Label1.Location = New System.Drawing.Point(30, 21)
            Label1.Name = "Label1"
            Label1.Size = New System.Drawing.Size(29, 12)
            Label1.TabIndex = 0
            Label1.Text = "主筋材料"
            '
            'ComboBox1
            '
            ComboBox1.FormattingEnabled = True
            ComboBox1.Location = New System.Drawing.Point(98, 18)
            ComboBox1.Name = "ComboBox1"
            ComboBox1.Size = New System.Drawing.Size(87, 20)
            ComboBox1.TabIndex = 1
            '
            'Label2
            '
            Label2.AutoSize = True
            Label2.Location = New System.Drawing.Point(30, 59)
            Label2.Name = "Label2"
            Label2.Size = New System.Drawing.Size(29, 12)
            Label2.TabIndex = 2
            Label2.Text = "箍筋材料"
            '
            'ComboBox2
            '
            ComboBox2.FormattingEnabled = True
            ComboBox2.Location = New System.Drawing.Point(98, 56)
            ComboBox2.Name = "ComboBox2"
            ComboBox2.Size = New System.Drawing.Size(87, 20)
            ComboBox2.TabIndex = 3
            '
            'Label3
            '
            Label3.AutoSize = True
            Label3.Location = New System.Drawing.Point(12, 106)
            Label3.Name = "Label3"
            Label3.Size = New System.Drawing.Size(73, 12)
            Label3.TabIndex = 4
            Label3.Text = "上保護層(外緣到主筋中心)"
            '
            'Label4
            '
            Label4.AutoSize = True
            Label4.Location = New System.Drawing.Point(12, 139)
            Label4.Name = "Label4"
            Label4.Size = New System.Drawing.Size(73, 12)
            Label4.TabIndex = 5
            Label4.Text = "下保護層(外緣到主筋中心)"
            '
            'TextBox1
            '
            TextBox1.Location = New System.Drawing.Point(148, 103)
            TextBox1.Name = "TextBox1"
            TextBox1.Size = New System.Drawing.Size(57, 22)
            TextBox1.TabIndex = 6
            '
            'TextBox2
            '
            TextBox2.Location = New System.Drawing.Point(148, 136)
            TextBox2.Name = "TextBox2"
            TextBox2.Size = New System.Drawing.Size(57, 22)
            TextBox2.TabIndex = 7
            '
            'Label5
            '
            Label5.AutoSize = True
            Label5.Location = New System.Drawing.Point(201, 106)
            Label5.Name = "Label5"
            Label5.Size = New System.Drawing.Size(37, 12)
            Label5.TabIndex = 8
            Label5.Text = "公分"
            '
            'Label6
            '
            Label6.AutoSize = True
            Label6.Location = New System.Drawing.Point(201, 139)
            Label6.Name = "Label6"
            Label6.Size = New System.Drawing.Size(37, 12)
            Label6.TabIndex = 9
            Label6.Text = "公分"
            '
            'Button1
            '
            Button1.Location = New System.Drawing.Point(287, 128)
            Button1.Name = "Button1"
            Button1.Size = New System.Drawing.Size(69, 30)
            Button1.TabIndex = 10
            Button1.Text = "確定"
            Button1.UseVisualStyleBackColor = True

            'Label7
            '
            Label7.AutoSize = True
            Label7.Location = New System.Drawing.Point(241, 42)
            Label7.Name = "Label6"
            Label7.Size = New System.Drawing.Size(37, 12)
            Label7.TabIndex = 11
            Label7.Text = "箍筋尺寸"

            'ComboBox3
            '
            Combobox3.FormattingEnabled = True
            Combobox3.Location = New System.Drawing.Point(200, 56)
            Combobox3.Name = "ComboBox3"
            Combobox3.Size = New System.Drawing.Size(87, 20)
            Combobox3.TabIndex = 12




            'inputF
            '
            InputForm.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            InputForm.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            InputForm.ClientSize = New System.Drawing.Size(368, 192)
            InputForm.Controls.Add(Button1)
            InputForm.Controls.Add(Label6)
            InputForm.Controls.Add(Label5)
            InputForm.Controls.Add(TextBox2)
            InputForm.Controls.Add(TextBox1)
            InputForm.Controls.Add(Label4)
            InputForm.Controls.Add(Label3)
            InputForm.Controls.Add(ComboBox2)
            InputForm.Controls.Add(Label2)
            InputForm.Controls.Add(ComboBox1)
            InputForm.Controls.Add(Label1)
            InputForm.Controls.Add(Combobox3)
            InputForm.Controls.Add(Label7)


            InputForm.Name = "inputF"
            InputForm.Text = "Input Data"
            Button1.DialogResult = DialogResult.OK
            InputForm.ResumeLayout(False)
            InputForm.PerformLayout()

            AddHandler Button1.Click, AddressOf myButtonClick


            Dim NumberNames As Integer
            Dim MyName() As String

            ret = SapModel.PropMaterial.GetNameList(NumberNames, MyName, eMatType.Rebar)


            If NumberNames = 0 Then
                MsgBox("No Rebar Data in SAP Model" & vbCrLf & "End Program")
                InputForm.Dispose()
                GoTo endsub
            Else
                For i = 0 To NumberNames - 1
                    ComboBox1.Items.Add(MyName(i))
                    ComboBox2.Items.Add(MyName(i))
                Next
                ComboBox1.SelectedIndex = 0
                ComboBox2.SelectedIndex = 0

            End If


            Combobox3.Items.Add("D10")
            Combobox3.Items.Add("D13")
            Combobox3.Items.Add("D16")
            ComboBox2.SelectedIndex = 1


            TextBox1.Text = "7.0"
            TextBox2.Text = "7.0"



            InputForm.ShowDialog()

            Dim MainBar As String
            Dim SecondaryBar As String

            If InputForm.DialogResult = DialogResult.OK Then

                MainBar = ComboBox1.SelectedItem
                SecondaryBar = ComboBox2.SelectedItem
                CoverTop = TextBox1.Text
                CoverBot = TextBox2.Text

                InputForm.Dispose()

            End If



            '=======================================================



            Dim StationNumber As Integer = 11


            Dim FileName As String

            Dim MatProp As String
            Dim t3 As Double
            Dim t2 As Double
            Dim Color As Long
            Dim Notes As String
            Dim GUID As String
            Dim MainBarSize As String
            Dim StrripSize As String
            Dim SplitIndex As Integer

            ret = SapModel.SetModelIsLocked(False)

            For i = 4 To LineCounter - 1 Step StationNumber
                SectName = xlSheet.Cells(i, 2).value.ToString

                FTopAreaStart = xlSheet.Cells(i + 1, 37).value
                FTopAreaEnd = xlSheet.Cells(i + 9, 37).value
                FBotAreaStart = xlSheet.Cells(i + 1, 39).value
                FBotAreaEnd = xlSheet.Cells(i + 9, 39).value


                SplitIndex = xlSheet.Cells(i + 1, 36).value.ToString.IndexOf("-")
                MainBarSize = xlSheet.Cells(i + 1, 36).value.ToString.Substring(SplitIndex + 1, 3)
                'Frame Note Data
                ret = SapModel.PropFrame.GetRectangle(SectName, FileName, MatProp, t3, t2, Color, Notes, GUID)
                ret = SapModel.PropFrame.SetRectangle(SectName, MatProp, t3, t2, , MainBarSize)



                ret = SapModel.PropFrame.SetRebarBeam(SectName, MainBar, SecondaryBar, CoverTop, CoverBot, FTopAreaStart, FTopAreaEnd, FBotAreaStart, FBotAreaEnd)





                If ret = 1 Then
                    MsgBox("Err11")
                    GoTo endsub2
                End If
            Next



            MsgBox("Finish !")
endsub:
            'xlBook.Save()

endsub2:
            xlBook.Close()
            xlApp.Quit()
            ISapPlugin.Finish(0)
        End Sub


        Private Sub myButtonClick(ByVal sender As Object, ByVal e As System.EventArgs)



        End Sub



    End Class
End Namespace

Namespace Rebar_Tool
    Public Class cPlugin

        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

            Dim p As Process
            p = Process.Start("v:\54833\SAP2000_RCRebar.exe")
            p.WaitForExit()



            ISapPlugin.Finish(0)
        End Sub

    End Class
End Namespace

Namespace DefineSecType
    Public Class cPlugin

        Public elev() As Double
        Public Lcount As Integer
        Public hasELzero As Boolean = False
        Public ConnectiveJoint() As Integer
        Public ConnCount As Integer

        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

            'If System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower <> "ctci.com.tw" Then
            '    MsgBox("CTCI Only")
            '    GoTo EndSub
            'End If

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


            Dim RebarSize_Girder As String
            Dim RebarSize_Beam As String
            Dim RebarSize_Torsion As String

            Dim RebarArea_Pri As Double
            Dim RebarArea_Sec As Double
            Dim RebarArea_Tor As Double
            Dim RebarArea As Double



            RebarSize_Girder = InputBox("Rebar Size - Girder " & vbCrLf & "EX : D25 / #8  etc...", "Input Rebar Size - Girder", "D25").ToUpper

            RebarSize_Beam = InputBox("Rebar Size - Beam  " & vbCrLf & "EX : D25 / #8 etc...", "Input Rebar Size - Beam", "D25").ToUpper
            RebarSize_Torsion = InputBox("Rebar Size - Torsion  " & vbCrLf & "EX : D19 / #6  etc...", "Input Rebar Size - Torsion", "D19").ToUpper

            RebarArea_Pri = GetRebarArea(RebarSize_Girder)
            If RebarArea_Pri = -1 Then MsgBox("Can't find " & RebarSize_Girder & " please contact RD Team")
            RebarArea_Sec = GetRebarArea(RebarSize_Beam)
            If RebarArea_Sec = -1 Then MsgBox("Can't find " & RebarSize_Beam & " please contact RD Team")
            RebarArea_Tor = GetRebarArea(RebarSize_Torsion)
            If RebarArea_Tor = -1 Then MsgBox("Can't find " & RebarSize_Torsion & " please contact RD Team")

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

            Dim BeamDataCol(NumberNames - 1, 12) As String       '只存桿件資料    是否需改用物件存資料?
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
                   , ErrorSummary, WarningSummary)           'Rebar Area  (R1 ~ R7)  這API有問題，讀Column ret仍然回傳0
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

            '==============設定Output Station 為  11 
            ret = SapModel.SelectObj.All
            ret = SapModel.FrameObj.SetOutputStations("ALL", 2, 0, 11, True, True, eItemType.SelectedObjects)
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

                    ret = SapModel.PropFrame.SetRebarBeam(NewSectionName, MatPropLong, MatPropConfine, CoverTop, CoverBot, AreaTL, AreaTR, AreaBL, AreaBR)
                    ret = SapModel.FrameObj.SetSection(BeamDataCol(i, 1), NewSectionName)

                End If
            Next






            MsgBox("Finish", , "Member Sort")
EndSub:


            ISapPlugin.Finish(0)
        End Sub

        Public Structure Seg_RebarArea
            Public TopLeft As Double
            Public TopMiddle As Double
            Public TopRight As Double
            Public BotLeft As Double
            Public BotMiddle As Double
            Public BotRight As Double
            Public Torsionbar As Double
        End Structure

        Public Function FindMaxRebarArea(ByVal TopArea() As Double, ByVal BotArea() As Double, ByVal TorRebar() As Double, ByVal Rebar As Seg_RebarArea) As Object
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


        Private Function MemberMark(ByRef Sx As Double, ByRef Sy As Double, ByRef Sz As Double, ByRef Ex As Double, ByRef Ey As Double, ByRef Ez As Double, ByRef P1 As String, ByRef P2 As String) As String

            '輸入起點/終點座標 透過規則定出該桿件類別

            Dim Lev As String

            If Sz = Ez Then
                For i = 0 To Lcount - 1
                    If Sz = elev(i) Then
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

            Dim BeamGirderDet As Boolean = False

            For i = 0 To ConnectiveJoint.Count - 2
                If P1 = ConnectiveJoint(i) Then
                    BeamGirderDet = True
                    Exit For
                End If
            Next

            If BeamGirderDet = True And Sz = Ez Then
                Lev = Lev + "G"
            ElseIf BeamGirderDet = False And Sz = Ez Then
                Lev = Lev + "B"
            End If


            If Sz <> Ez And Sx = Ex And Sy = Ey Then
                Lev = Lev + "C"
            Else
                If Math.Round(Sy, 2) = Math.Round(Ey, 2) Then Lev = Lev + "X"
                If Math.Round(Sx, 2) = Math.Round(Ex, 2) Then Lev = Lev + "Y"

            End If

            Return Lev


        End Function




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
                    hasELzero = True
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

    End Class
End Namespace




Namespace OutputRCResult

    Public Class cPlugin



        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)


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
            StirrupSize = InputBox("輸入箍筋尺寸   EX :  D10 / #3").ToUpper
            SideBarSize = InputBox("輸入腹筋尺寸   EX :  D19 / #6").ToUpper

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

            '===========

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
            '================================
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
            '==================================
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
            '===================================
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





        'Public Function GetRebarArea(ByRef BarSize As String) As Double

        '    Select Case BarSize
        '        Case "D6"
        '            GetRebarArea = 0.3167
        '        Case "D10"
        '            GetRebarArea = 0.7133
        '        Case "D13"
        '            GetRebarArea = 1.267
        '        Case "D16"
        '            GetRebarArea = 1.986
        '        Case "D19"
        '            GetRebarArea = 2.865
        '        Case "D22"
        '            GetRebarArea = 3.871
        '        Case "D25"
        '            GetRebarArea = 5.067
        '        Case "D29"
        '            GetRebarArea = 6.469
        '        Case "D32"
        '            GetRebarArea = 8.143
        '        Case "D36"
        '            GetRebarArea = 10.07
        '        Case "D39"
        '            GetRebarArea = 12.19
        '        Case "D43"
        '            GetRebarArea = 14.52
        '        Case "D50"
        '            GetRebarArea = 19.79
        '        Case "D57"
        '            GetRebarArea = 25.79
        '    End Select


        'End Function

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


    End Class

End Namespace







Namespace OpenManual

    Public Class cPlugin

        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)

            Dim ManualForm As New Manual

            ManualForm.ShowDialog()
            ISapPlugin.Finish(0)
        End Sub

    End Class
End Namespace

'主程式已嵌入Menu form中
Namespace FunctionMenu

    Public Class cPlugin


        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)


            If frmMenu Is Nothing Then
                frmMenu = New Menu
                frmMenu.SapModel = SapModel
                frmMenu.ISapPlugin = ISapPlugin


                Dim OSVer As String
                OSVer = System.Environment.OSVersion.VersionString
                If OSVer.Contains("6.1") Then
                    SE.SkinFile = "C:\Program Files (x86)\Computers and Structures\SAP2000 16\CTCIPlugin\ssk\MSN.ssk"
                Else
                    SE.SkinFile = "C:\Program Files\Computers and Structures\SAP2000 16\CTCIPlugin\ssk\MSN.ssk"
                End If
                frmMenu.SkinEngine1 = SE

            End If

            frmMenu.Focus()
            frmMenu.KeyPreview = True
            frmMenu.Show()


            SE.Active = True

            ISapPlugin.Finish(0)
        End Sub

    End Class
End Namespace


'Namespace VerCheck
'    Public Class cPlugin
'        Public Sub Main(ByRef SapModel As SAP2000v20.cSapModel, ByRef ISapPlugin As SAP2000v20.cPluginCallback)
'            Dim manualForm As New Manual
'            manualForm.ShowDialog()
'            MsgBox("This is SAP 2000 V15.1.0 Plug-in  (v1.20)", MsgBoxStyle.Information)  '每次需修改

'            Dim UsedVer As String
'            UsedVer = "1.20"   '每次需修改

'            FileOpen(100, "N:\SAP2000_Plugin\V15.1.0\VersionCheck.txt", OpenMode.Input)
'            Dim CurrentVersion As String
'            CurrentVersion = LineInput(100)

'            If UsedVer <> CurrentVersion Then
'                MsgBox("Your Version :  " & UsedVer & vbCrLf & "Current Version : " & CurrentVersion & vbCrLf & vbCrLf & _
'                        "Please Update SAP2000 Plug-in", MsgBoxStyle.Exclamation)
'            Else
'                MsgBox("Plug-in Version Chekck   :   OK! ")
'            End If

'            FileClose(100)

'            ISapPlugin.Finish(0)
'        End Sub

'    End Class
'End Namespace


