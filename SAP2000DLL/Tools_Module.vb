'Imports System.Runtime.InteropServices


Module Tools_Module

    Public Pool_Length As Double
    Public Pool_Width As Double
    Public Pool_Height As Double
    Public Divide_X As Integer
    Public Divide_Y As Integer
    Public Divide_Z As Integer
    Public WallThk As Double
    Public BottomSlabThk As Double
    Public WithTopSlab As Boolean
    Public TopSlabThk As Double
    Public TopWallThk As Double
    Public BottomWallThk As Double
    Public WaterPressure_Top As Double
    Public WaterPressure_Bottom As Double
    Public WaterPressure2_Top As Double
    Public WaterPressure2_Bottom As Double

    Public WaterPressure3_Top As Double
    Public WaterPressure3_Bottom As Double

    Public WaterPressure4_Top As Double
    Public WaterPressure4_Bottom As Double
    Public WaterDepth As Double

    Public VerSoil_K As Double
    Public HoriSoil_K As Double
    Public VerSeismicCoeff As Double


    Public CombList() As String
    Public ReportFileName As String
    Public DeflectionCriteria As Integer
    Public SelectCombFlag As Boolean = False
    Public JointSequenceFile As String
    Public OpenFileErrorFlag As Boolean = False
    Public DefChkCriteria As Boolean = True
    Public ShowNodeList As Boolean = True
    Public ShowCriNode As Boolean = True
    Public ShowSectName As Boolean = True
    Public CheckDeflectionRelative As Boolean = True
    Public CheckDeflectionAbsolute As Boolean = False

    Public RebarSize_Girder As String
    Public RebarSize_Beam As String
    Public RebarSize_Torsion As String
    Public RebarArea_Pri As Double
    Public RebarArea_Sec As Double
    Public RebarArea_Tor As Double
    Public RebarArea As Double

    'Public Sub code9_1_1_2(ByRef AA As String, ByRef BB As String, ByRef CC As String)

    '    AA = "aaaaa"

    '    BB = "BBBBB"

    '    CC = AA + BB

    'End Sub


    Public Function CheckDomain() As Boolean

        Dim checkCTCIGroup As Boolean = False

        If System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "ctci.com.tw" Then
            checkCTCIGroup = True
        ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "jdec.com.cn" Then
            checkCTCIGroup = True
        ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "cimas.com.vn" Then
            checkCTCIGroup = True
        ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "cinda.in" Then
            checkCTCIGroup = True
        ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "ctci.co.th" Then
            checkCTCIGroup = True
        ElseIf System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties.DomainName.ToString.ToLower = "ctcim.com.tw" Then
            checkCTCIGroup = True
        End If


        CheckDomain = checkCTCIGroup

    End Function

    Public Function GetRebarArea(ByRef BarSize As String) As Double

        'Area unit is Square CM

        Select Case BarSize
            Case "D6", "#2"
                GetRebarArea = 0.3167
            Case "D10", "#3"
                GetRebarArea = 0.7133
            Case "D12"
                GetRebarArea = 1.131
            Case "D13", "#4"
                GetRebarArea = 1.267
            Case "D16", "#5"
                GetRebarArea = 1.986
            Case "D19", "#6"
                GetRebarArea = 2.865
            Case "D20"
                GetRebarArea = 3.142
            Case "D22", "#7"
                GetRebarArea = 3.871
            Case "D25", "#8"
                GetRebarArea = 5.067
            Case "D29", "#9"
                GetRebarArea = 6.469
            Case "D32", "#10"
                GetRebarArea = 8.143
            Case "D36", "#11"
                GetRebarArea = 10.07
            Case "D39", "#12"
                GetRebarArea = 12.19
            Case "D43", "#13"
                GetRebarArea = 14.52
            Case "D50", "#14"
                GetRebarArea = 19.79
            Case "D57", "#15"
                GetRebarArea = 25.79
            Case "T10"
                GetRebarArea = 0.785
            Case "T12"
                GetRebarArea = 1.13
            Case "T16"
                GetRebarArea = 2.01
            Case "T20"
                GetRebarArea = 3.14
            Case "T25"
                GetRebarArea = 4.91
            Case "T32"
                GetRebarArea = 8.04

            Case Else
                GetRebarArea = -1

        End Select


    End Function



End Module
