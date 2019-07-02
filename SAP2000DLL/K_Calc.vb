Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Media.Media3D
Imports SAP2000v20
Imports System.IO


Public Class K_Calc

    Public SapModel As SAP2000v20.cSapModel
    Public ISapPlugin As SAP2000v20.cPluginCallback
    Public MainI, MainLength, GA, GB As Double


   

    Private Sub K_Calc_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnSelectMain_Click(sender As Object, e As EventArgs) Handles btnSelectMain.Click
        Dim ret As Long
        Dim NumberNames As Long
        Dim GroupName() As String


        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        Dim PropName As String
        Dim SAuto As String

       
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

        Dim Point1 As String
        Dim Point2 As String

        Dim X1, Y1, Z1 As Double
        Dim X2, Y2, Z2 As Double

        If NumberItems > 1 Then
            MsgBox("One member only")
        Else
            txtBoxMain.Text = ObjectName(0)

            ret = SapModel.FrameObj.GetSection(ObjectName(0), PropName, SAuto)

            ret = SapModel.PropFrame.GetGeneral(PropName, FileName, MatProp, t3, t2, Area, as2, as3, Torsion, I22, I33, S22, S33, Z22, Z33, R22, R33, Color, Notes, GUID)

            If btnMajor.Checked = True Then
                MainI = I33
            Else
                MainI = I22
            End If

            ret = SapModel.FrameObj.GetPoints(ObjectName(0), Point1, Point2)

            ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)     'point1 coord
            ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)     'point2 coord

            MainLength = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2) ^ 0.5


        End If


    End Sub

    Private Sub btnSelectGASub_Click(sender As Object, e As EventArgs) Handles btnSelectGASub.Click

        Dim ret As Long
        Dim NumberItems As Integer
        Dim ObjectType() As Integer
        Dim ObjectName() As String
        ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)

        If NumberItems = 0 Then
            If btnHinge.Checked Then
                GA = 10
            End If
            If btnFixed.Checked Then
                GA = 1
            End If
        End If

        Dim eleList As String = ""

        If NumberItems > 0 Then

            For i = 0 To NumberItems
                eleList += ObjectName(i) & ","
            Next
            txtGAele.Text = eleList

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

            Dim PropName As String
            Dim SAuto As String

            Dim Point1 As String
            Dim Point2 As String

            Dim X1, Y1, Z1 As Double
            Dim X2, Y2, Z2 As Double

            Dim Length As Double

            For i = 0 To NumberItems - 1
                ret = SapModel.FrameObj.GetSection(ObjectName(0), PropName, SAuto)

                ret = SapModel.PropFrame.GetGeneral(PropName, FileName, MatProp, t3, t2, Area, as2, as3, Torsion, I22, I33, S22, S33, Z22, Z33, R22, R33, Color, Notes, GUID)

                If btnMajor.Checked = True Then
                    MainI = I33
                Else
                    MainI = I22
                End If

                ret = SapModel.FrameObj.GetPoints(ObjectName(0), Point1, Point2)

                ret = SapModel.PointObj.GetCoordCartesian(Point1, X1, Y1, Z1)     'point1 coord
                ret = SapModel.PointObj.GetCoordCartesian(Point2, X2, Y2, Z2)     'point2 coord

               





            Next


        End If



    End Sub
End Class