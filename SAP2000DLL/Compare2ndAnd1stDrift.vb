Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Media.Media3D
Imports SAP2000v20
Imports System.IO

Public Class Compare2ndAnd1stDrift

    Private Sub TextBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox1.MouseDoubleClick
        OpenFileDialog1.ShowDialog()
        TextBox1.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub TextBox2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox2.MouseDoubleClick
        OpenFileDialog2.ShowDialog()
        TextBox2.Text = OpenFileDialog2.FileName
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            FileOpen(10, OpenFileDialog1.FileName, OpenMode.Input)
            FileOpen(20, OpenFileDialog2.FileName, OpenMode.Input)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Compare 2nd Order & 1St Order Drift")
            FileClose()
            OpenFileErrorFlag = True
            Exit Sub
        End Try

        '===========處理資料

        Dim charseparators() As Char = " "

        Dim lineText As String
        Dim dummy As String
        Dim FirstOrderData() As String

        Dim DriftData_1 As New DriftData
        Dim DriftData_2 As New DriftData

        Dim Coll_FileA As List(Of DriftData)


        dummy = LineInput(10)

        Do Until EOF(10)
            lineText = LineInput(10)
            FirstOrderData = lineText.Split(charseparators, StringSplitOptions.RemoveEmptyEntries)
            DriftData_1.CL_Name = FirstOrderData(0)
            DriftData_1.Node = FirstOrderData(1)
            DriftData_1.H_200 = FirstOrderData(2)
            DriftData_1.Comb_X = FirstOrderData(3)
            DriftData_1.Tx1 = FirstOrderData(4)
            DriftData_1.Tx2 = FirstOrderData(5)
            DriftData_1.Dx = FirstOrderData(6)
            DriftData_1.X_Check = FirstOrderData(7)
            DriftData_1.Comb_Y = FirstOrderData(8)
            DriftData_1.Ty1 = FirstOrderData(9)
            DriftData_1.Ty2 = FirstOrderData(10)
            DriftData_1.Dy = FirstOrderData(11)
            DriftData_1.Y_Check = FirstOrderData(12)




        Loop



ExitSub:
        FileClose(10)
        FileClose(20)

    End Sub


    Structure DriftData
        Dim CL_Name As String
        Dim Node As String
        Dim H_200 As Double
        Dim Comb_X As String
        Dim Tx1, Tx2, Dx As Double
        Dim X_Check As String
        Dim Comb_Y As String
        Dim Ty1, Ty2, Dy As Double
        Dim Y_Check As String
    End Structure



    Private Sub Compare2ndAnd1stDrift_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class