'==============================================================================
' CST-LinXi-Macro  -  CST eigenmode slow-wave structure user-defined watch macro
'
' Copyright (c) 2026 She and Me
' SPDX-License-Identifier: MIT
'
' Released under the MIT License. Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation
' files, to deal in the Software without restriction. Full text: see LICENSE.
'
'==============================================================================

Option Explicit

Private Const CORE_AUTHOR As String = " Lin Xi and She & Me "

Private Const CLight   As Double = 299792458#
Private Const DegToRad As Double = Pi / 180#

Private Const WAIT_SEC As Double = 0.1

Private Const MAX_MODE_PARAMS   As Integer = 1000
Private Const MAX_PTS_PER_DIR   As Long = 1000
Private Const MAX_CROSS_POINTS  As Long = 200000
Private Const MAX_FREQ_RETRY    As Integer = 30
Private Const FREQ_RETRY_WAIT   As Double = 0.2

Private Const FREQ_FLAT_TOL     As Double = 1E-09
Private Const FREQ_FLAT_MAX     As Integer = 5
Private Const KC_SANITY_MAX     As Double = 1E+08

Private Const MAX_PATH_SAFE        As Integer = 250
Private Const FILE_TAG_FIXED       As Integer = 24

Private Enum LogLevel
    llDebug = 0
    llInfo = 1
    llWarning = 2
    llError = 3
    llCritical = 4
End Enum

Private Const LOG_TAG_DEBUG        As String = " DEBUG "
Private Const LOG_TAG_INFO         As String = " INFO  "
Private Const LOG_TAG_WARNING      As String = " WARN  "
Private Const LOG_TAG_ERROR        As String = " ERROR "
Private Const LOG_TAG_CRITICAL     As String = " CRIT! "

Private Const PFX_MSG_INFO         As String = "MacroMsg [INFO] "
Private Const PFX_MSG_WARNING      As String = "MacroMsg [WARN] "
Private Const PFX_MSG_ERROR        As String = "MacroMsg [ERROR] "
Private Const PFX_MSG_CRITICAL     As String = "MacroMsg [CRIT] "

Private Const LOG_NAME             As String = "_macro_log.txt"
Private Const LOG_MSG_MAX_LEN      As Long = 10000
Private Const LOG_ROTATE_BYTES     As Long = 5242880
Private Const MSG_ABORT_MACRO      As String = vbCrLf & vbCrLf & "A fatal error was found! The macro has aborted the parameter sweep!" & vbCrLf & _
                                               "Check the parameter settings before restarting the sweep or running the macro again!" & vbCrLf & _
                                               "Error message: "

Private Const F_BETA    As String = "beta"
Private Const F_ZPIERCE As String = "ZpierceAvg"
Private Const F_VPHASE  As String = "vphase"
Private Const F_PHASE   As String = "phase"
Private Const F_RESULT_GROUP   As String = "User-Defined Macro Result"

Private Const CFG_FILE_NAME    As String = "LinXi.ini"
Private Const CFG_SECTION_VER  As String = "VERSION"

Private Const PW_SRC_FROM_EH  As Integer = 1
Private Const PW_SRC_FROM_CST As Integer = 2
Private Const PW_SRC_INVALID  As Integer = 0

Private Const REGION_BOTH     As Integer = 0
Private Const REGION_FORWARD  As Integer = 1
Private Const REGION_BACKWARD As Integer = 2

Private Const ENABLE_FLAG_SCALE  As Long = 10
Private Const ENABLE_FLAG_EPS    As Double = 1E-06

Private g_bIsTetra      As Boolean
Private g_bFirstFieldLog As Boolean
Private g_bAllModes     As Boolean
Private g_abModeFlag()  As Boolean
Private g_bGridFailed   As Boolean
Private g_abPhaseDead() As Boolean
Private g_anPhaseFlat() As Long
Private g_abPhaseWarned() As Boolean
Private g_nPhaseStateModes As Long

Private g_sPointsFile   As String
Private g_sLastGroupKey As String
Private g_sUnit         As String
Private g_sDirLabel    As String
Private g_sGroupPath   As String
Private g_sGroupKey    As String
Private g_sLogFile      As String
Private g_sTemp         As String
Private g_sResult       As String
Private g_sFileTag      As String
Private g_nFileTagBudget As Long

Private g_bLogInited        As Boolean
Private g_nLogFileMinLevel  As LogLevel
Private g_sLogLastError     As String

Private g_sVersionString As String
Private g_sVersionLabel  As String
Private g_sAuthorName    As String
Private g_sReleaseDate   As String
Private g_sPlatform      As String
Private g_sCstVerMin     As String
Private g_sCstVerMax     As String
Private g_sCfgPath       As String
Private g_sCfgNote       As String
Private g_sCfgVerRaw     As String
Private g_sCfgLblRaw     As String
Private g_sCfgDllMinRaw  As String
Private g_nDllMinVer     As Long
Private g_bCfgOK         As Boolean

Private g_iDir              As Integer
Private g_iNumModes         As Integer
Private g_aSelectedModes()  As Integer
Private g_nSelectedModes    As Integer
Private g_iPowerFlowSrc     As Integer
Private g_iRegionMode       As Integer

Private g_aExRe() As Double, g_aExIm() As Double
Private g_aEyRe() As Double, g_aEyIm() As Double
Private g_aEzRe() As Double, g_aEzIm() As Double
Private g_aHxRe() As Double, g_aHxIm() As Double
Private g_aHyRe() As Double, g_aHyIm() As Double
Private g_aHzRe() As Double, g_aHzIm() As Double
Private g_aPxRe() As Double, g_aPxIm() As Double
Private g_aPyRe() As Double, g_aPyIm() As Double
Private g_aPzRe() As Double, g_aPzIm() As Double

Private g_aBufBeta()    As String
Private g_aBufZpierce() As String
Private g_aBufPhase()   As String
Private g_nBufModes     As Long

Private g_arrDw1() As Double, g_arrDw2() As Double
Private g_arrMidL() As Double, g_arrDwL() As Double
Private g_cMin(1 To 3) As Double, g_cMax(1 To 3) As Double
Private g_xRef As Double, g_yRef As Double, g_zRef As Double
Private g_dPitchSI   As Double
Private g_dPitchUnit As Double
Private g_dUnitToSI  As Double

Private g_n1 As Long
Private g_n2 As Long
Private g_nCross As Long
Private g_nL As Long

Declare Function CCoreVersion Lib "LinXi.dll" (ByVal author As String) As Long

Declare Function CParseFieldFile Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal sPath As String, ByVal nExpected As Long, _
    ByRef firstX As Double, ByRef firstY As Double, ByRef firstZ As Double, _
    ByRef exRe As Double, ByRef exIm As Double, _
    ByRef eyRe As Double, ByRef eyIm As Double, _
    ByRef ezRe As Double, ByRef ezIm As Double) As Long

Declare Function CParsePowerFlowFile Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal sPath As String, ByVal nExpected As Long, _
    ByRef pxRe As Double, ByRef pxIm As Double, _
    ByRef pyRe As Double, ByRef pyIm As Double, _
    ByRef pzRe As Double, ByRef pzIm As Double) As Long

Declare Function CComputePowerFlow Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal iDir As Long, ByVal n1 As Long, ByVal n2 As Long, _
    ByRef dw1 As Double, ByRef dw2 As Double, _
    ByRef exRe As Double, ByRef exIm As Double, _
    ByRef eyRe As Double, ByRef eyIm As Double, _
    ByRef ezRe As Double, ByRef ezIm As Double, _
    ByRef hxRe As Double, ByRef hxIm As Double, _
    ByRef hyRe As Double, ByRef hyIm As Double, _
    ByRef hzRe As Double, ByRef hzIm As Double) As Double

Declare Function CIntegrateLongitudinal Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal iDir As Long, ByVal nL As Long, ByVal nCross As Long, _
    ByVal beta As Double, ByVal unitToSI As Double, ByVal pitchSI As Double, _
    ByRef midL As Double, ByRef dwL As Double, _
    ByRef exRe As Double, ByRef exIm As Double, _
    ByRef eyRe As Double, ByRef eyIm As Double, _
    ByRef ezRe As Double, ByRef ezIm As Double, _
    ByRef vre As Double, ByRef vim As Double, ByRef eabs As Double) As Long

Declare Function CIntegratePowerFlow Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal iDir As Long, ByVal n1 As Long, ByVal n2 As Long, _
    ByRef dw1 As Double, ByRef dw2 As Double, _
    ByRef pxRe As Double, ByRef pxIm As Double, _
    ByRef pyRe As Double, ByRef pyIm As Double, _
    ByRef pzRe As Double, ByRef pzIm As Double) As Double

Declare Function CComputeTetraStep Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal len1 As Double, ByVal len2 As Double, ByVal lenL As Double, _
    ByVal minEdge As Double, ByVal maxEdge As Double, _
    ByVal maxCrossPoints As Long, ByVal maxPtsPerDir As Long, _
    ByRef n1 As Long, ByRef n2 As Long, ByRef nL As Long, _
    ByRef h As Double) As Long

Declare Function CBuildUniformGrid Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal n1 As Long, ByVal n2 As Long, ByVal nL As Long, _
    ByVal cMin1 As Double, ByVal cMax1 As Double, _
    ByVal cMin2 As Double, ByVal cMax2 As Double, _
    ByVal cMinL As Double, ByVal cMaxL As Double, _
    ByVal unitToSI As Double, _
    ByRef mid1 As Double, ByRef dw1 As Double, _
    ByRef mid2 As Double, ByRef dw2 As Double, _
    ByRef midL As Double, ByRef dwL As Double, _
    ByVal iDir As Long, _
    ByVal xRef As Double, ByVal yRef As Double, ByVal zRef As Double, _
    ByVal pointsFile As String) As Long

Declare Function CBuildHexaGrid Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal iDir As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal nL As Long, _
    ByRef coord1 As Double, ByRef coord2 As Double, ByRef coordL As Double, _
    ByVal unitToSI As Double, _
    ByVal xRef As Double, ByVal yRef As Double, ByVal zRef As Double, _
    ByVal pointsFile As String, _
    ByRef mid1 As Double, ByRef dw1 As Double, _
    ByRef mid2 As Double, ByRef dw2 As Double, _
    ByRef midL As Double, ByRef dwL As Double) As Long

Declare Function CBuildCurveFiles Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal sTempPrefix As String, ByVal sSuffix As String, _
    ByVal iMode As Long, ByVal cLight As Double, _
    ByVal sOutBeta As String, ByVal sOutPhase As String, _
    ByVal sOutVp As String, ByVal sOutKc As String, _
    ByRef nBetaPt As Long, ByRef nPhasePt As Long, _
    ByRef nVpPt As Long, ByRef nKcPt As Long) As Long

Private Sub ParameterSweepWatch(ByVal action As Integer)

    PreConfiguration action

    Select Case action
        Case 0 : InitializationPhase
        Case 1 : ProcessingPhase
        Case 2 : FinalizationPhase
    End Select
End Sub

Private Sub PreConfiguration(ByVal action As Integer)
    Dim sMeshType As String

    LoadMacroConfig

    sMeshType = Mesh.GetMeshType

    g_bIsTetra    = (sMeshType = "Tetrahedral")
    g_sTemp       = GetProjectPath("Temp")
    g_sResult     = GetProjectPath("Result")
    g_dUnitToSI   = Units.GetGeometryUnitToSI
    g_sUnit       = Units.GetGeometryUnit
    g_sLogFile    = g_sTemp & LOG_NAME
    g_sPointsFile = g_sTemp & "_field_points.txt"
    g_iNumModes   = Solver.AKSGetNumberOfModes

    g_nFileTagBudget = MAX_PATH_SAFE - Len(g_sTemp) - FILE_TAG_FIXED

    If action = 0 Then InitLogFile

    If Not g_bCfgOK Then AbortOnConfigError

    Dim Macro_flag As Double
    Dim bNeedSetupDialog As Boolean
    Dim iFlagSrc As Integer
    Dim iFlagRegion As Integer

    bNeedSetupDialog = False
    g_iRegionMode = REGION_BOTH

    If DoesParameterExist("Macro_SweepWatch_Enable") Then
        Macro_flag = RestoreParameter("Macro_SweepWatch_Enable")

        If Abs(Macro_flag) < 1E-09 Then
            SetParameterDescription("Macro_SweepWatch_Enable", "Macro disabled state")
            End
        ElseIf DecodeEnableFlag(Macro_flag, iFlagSrc, iFlagRegion) Then
            g_iPowerFlowSrc = iFlagSrc
            g_iRegionMode = iFlagRegion
            SetParameterDescription "Macro_SweepWatch_Enable", EnableFlagText(iFlagSrc, iFlagRegion)
        Else
            bNeedSetupDialog = True
            g_iPowerFlowSrc = 0
            g_iRegionMode = REGION_BOTH
            LogWarning "  Macro_SweepWatch_Enable = " & Format(Macro_flag, "0.0####") & _
                " is not a valid enable-flag value; the parameter setup dialog will be shown again for this run"
        End If
    Else
        bNeedSetupDialog = True
        g_iPowerFlowSrc = 0
        g_iRegionMode = REGION_BOTH
    End If

    With Boundary
        Dim nPeriodic As Integer
        nPeriodic = 0
        If .GetXmin = "periodic" Then g_iDir = 1 : nPeriodic = nPeriodic + 1
        If .GetYmin = "periodic" Then g_iDir = 2 : nPeriodic = nPeriodic + 1
        If .GetZmin = "periodic" Then g_iDir = 3 : nPeriodic = nPeriodic + 1
        If nPeriodic > 1 Then
            LogCritical("Multiple periodic directions detected!", True)
            End
        ElseIf g_iDir = 0 Then
            LogCritical("No periodic boundary condition is defined!", True)
            End
        End If
        .GetCalculationBox g_cMin(1), g_cMax(1), _
            g_cMin(2), g_cMax(2), _
            g_cMin(3), g_cMax(3)
        g_dPitchUnit = g_cMax(g_iDir) - g_cMin(g_iDir)
        If g_dPitchUnit <= 0 Then
            LogCritical("Period length is zero; the mesh may be too coarse", True)
            End
        End If
        g_dPitchSI = g_dPitchUnit * g_dUnitToSI
    End With

    Dim labels(1 To 3) As String
    labels(1) = "X": labels(2) = "Y": labels(3) = "Z"
    g_sDirLabel = labels(g_iDir)

    Dim pos(1 To 3) As Double
    pos(1) = 0#: pos(2) = 0#: pos(3) = 0#

    Dim bRefPosExist As Boolean
    Dim bAnyModeParamExists As Boolean
    Dim iChk As Integer

    bRefPosExist = DoesParameterExist("Kc_RefPos_x") And _
        DoesParameterExist("Kc_RefPos_y") And _
        DoesParameterExist("Kc_RefPos_z")

    bAnyModeParamExists = False
    For iChk = 1 To g_iNumModes
        If DoesParameterExist("Macro_Mode" & CStr(iChk)) Then
            bAnyModeParamExists = True
            Exit For
        End If
    Next iChk

    If bRefPosExist And bAnyModeParamExists And (Not bNeedSetupDialog) Then
        pos(1) = RestoreParameter("Kc_RefPos_x")
        pos(2) = RestoreParameter("Kc_RefPos_y")
        pos(3) = RestoreParameter("Kc_RefPos_z")

        g_nSelectedModes = 0
        ReDim g_aSelectedModes(0 To g_iNumModes - 1)
        For iChk = 1 To g_iNumModes
            If DoesParameterExist("Macro_Mode" & CStr(iChk)) Then
                g_aSelectedModes(g_nSelectedModes) = iChk
                g_nSelectedModes = g_nSelectedModes + 1
                SetParameterDescription "Macro_Mode" & CStr(iChk), _
                    "Mode " & CStr(iChk) & " interaction impedance calculation flag"
            End If
        Next iChk

        If g_nSelectedModes = 0 Then
            g_bAllModes = True
        ElseIf g_nSelectedModes = g_iNumModes Then
            g_bAllModes = True
            g_nSelectedModes = 0
        Else
            g_bAllModes = False
            ReDim Preserve g_aSelectedModes(0 To g_nSelectedModes - 1)
        End If

        Dim iIdx As Integer
    Else

        If DoesParameterExist("Kc_RefPos_x") Then
            pos(1) = RestoreParameter("Kc_RefPos_x")
        End If
        If DoesParameterExist("Kc_RefPos_y") Then
            pos(2) = RestoreParameter("Kc_RefPos_y")
        End If
        If DoesParameterExist("Kc_RefPos_z") Then
            pos(3) = RestoreParameter("Kc_RefPos_z")
        End If

        If Not ShowParamsDialog(pos, g_bAllModes) Then
            StoreParameter("Macro_SweepWatch_Enable", 0#)
            SetParameterDescription("Macro_SweepWatch_Enable", "Macro disabled state")
            LogInfo "User cancelled the macro setup! The enable flag is set to 0, the macro will not run, but the CST sweep will continue!"
            End
        End If

        StoreParameter "Macro_SweepWatch_Enable", _
            CDbl(g_iPowerFlowSrc + 1) + CDbl(g_iRegionMode) / 10#
        SetParameterDescription "Macro_SweepWatch_Enable", _
            EnableFlagText(g_iPowerFlowSrc, g_iRegionMode)

        pos(g_iDir) = 0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir))
        Dim iPos As Integer
        For iPos = 1 To 3
            If Abs(pos(iPos)) < 1e-12 Then pos(iPos) = 0#
        Next iPos

        StoreParameter "Kc_RefPos_x", pos(1)
        StoreParameter "Kc_RefPos_y", pos(2)
        StoreParameter "Kc_RefPos_z", pos(3)
        SetParameterDescription "Kc_RefPos_x", "X coordinate of the interaction impedance reference point"
        SetParameterDescription "Kc_RefPos_y", "Y coordinate of the interaction impedance reference point"
        SetParameterDescription "Kc_RefPos_z", "Z coordinate of the interaction impedance reference point"

        On Error Resume Next
        For iChk = 1 To MAX_MODE_PARAMS
            If DoesParameterExist("Macro_Mode" & CStr(iChk)) Then
                DeleteParameter "Macro_Mode" & CStr(iChk)
            End If
        Next iChk
        On Error GoTo 0

        If g_bAllModes Then
            For iChk = 1 To g_iNumModes
                StoreParameter "Macro_Mode" & CStr(iChk), 0#
                SetParameterDescription "Macro_Mode" & CStr(iChk), _
                    "Mode " & CStr(iChk) & " interaction impedance calculation flag"
            Next iChk
        Else
            For iIdx = 0 To g_nSelectedModes - 1
                iChk = g_aSelectedModes(iIdx)
                StoreParameter "Macro_Mode" & CStr(iChk), 0#
                SetParameterDescription "Macro_Mode" & CStr(iChk), _
                    "Mode " & CStr(iChk) & " interaction impedance calculation flag"
            Next iIdx
        End If
    End If

    g_xRef = pos(1)
    g_yRef = pos(2)
    g_zRef = pos(3)

    Select Case g_sDirLabel
        Case "X"
            If g_yRef <= g_cmin(2) Or g_yRef >= g_cmax(2) Then LogCritical("Y coordinate of the interaction impedance reference point lies outside the computation domain!", True)
            If g_zRef <= g_cmin(3) Or g_zRef >= g_cmax(3) Then LogCritical("Z coordinate of the interaction impedance reference point lies outside the computation domain!", True)
        Case "Y"
            If g_xRef <= g_cmin(1) Or g_xRef >= g_cmax(1) Then LogCritical("X coordinate of the interaction impedance reference point lies outside the computation domain!", True)
            If g_zRef <= g_cmin(3) Or g_zRef >= g_cmax(3) Then LogCritical("Z coordinate of the interaction impedance reference point lies outside the computation domain!", True)
        Case "Z"
            If g_xRef <= g_cmin(1) Or g_xRef >= g_cmax(1) Then LogCritical("X coordinate of the interaction impedance reference point lies outside the computation domain!", True)
            If g_yRef <= g_cmin(2) Or g_yRef >= g_cmax(2) Then LogCritical("Y coordinate of the interaction impedance reference point lies outside the computation domain!", True)
    End Select

    BuildModeFlagTable
End Sub

Private Sub InitializationPhase()

    g_sLastGroupKey = ""
    On Error Resume Next
    Kill g_sTemp & "_groups.txt"
    On Error GoTo 0

    DeleteTreeItemRecursive "1D Results\" & F_RESULT_GROUP

    If g_iNumModes = 0 Then
        LogCritical "No eigenmode found", True
        End
    End If

    On Error Resume Next
    DeleteFilesByPattern g_sTemp & "_beta_*.sig"
    DeleteFilesByPattern g_sTemp & "_ZpierceAvg_*.sig"
    DeleteFilesByPattern g_sTemp & "_phase_*.sig"
    DeleteFilesByPattern g_sTemp & "_prevfreq_*.txt"
    DeleteFilesByPattern g_sTemp & "_E_mode*.txt"
    DeleteFilesByPattern g_sTemp & "_H_mode*.txt"
    DeleteFilesByPattern g_sTemp & "_P_mode*.txt"
    DeleteFilesByPattern g_sTemp & "_curve*.txt"
    Kill g_sPointsFile
    On Error GoTo 0

    g_nCross = 0
    g_bGridFailed = False
    g_nBufModes = 0
    g_nPhaseStateModes = 0

    Dim iCleanup As Integer
    For iCleanup = g_iNumModes + 1 To MAX_MODE_PARAMS
        On Error Resume Next
        If DoesParameterExist("Macro_Mode" & CStr(iCleanup)) Then
            DeleteParameter "Macro_Mode" & CStr(iCleanup)
        End If
        On Error GoTo 0
    Next iCleanup

    Select Case g_sDirLabel
        Case "X"
            StoreParameter "Kc_RefPos_x", IIf(0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)) < 1e-12, 0, 0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)))
            SetParameterDescription "Kc_RefPos_x", "X coordinate of the interaction impedance reference point, periodic direction"
            SetParameterDescription "Kc_RefPos_y", "Y coordinate of the interaction impedance reference point"
            SetParameterDescription "Kc_RefPos_z", "Z coordinate of the interaction impedance reference point"
        Case "Y"
            StoreParameter "Kc_RefPos_y", IIf(0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)) < 1e-12, 0, 0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)))
            SetParameterDescription "Kc_RefPos_x", "X coordinate of the interaction impedance reference point"
            SetParameterDescription "Kc_RefPos_y", "Y coordinate of the interaction impedance reference point, periodic direction"
            SetParameterDescription "Kc_RefPos_z", "Z coordinate of the interaction impedance reference point"
        Case "Z"
            StoreParameter "Kc_RefPos_z", IIf(0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)) < 1e-12, 0, 0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)))
            SetParameterDescription "Kc_RefPos_x", "X coordinate of the interaction impedance reference point"
            SetParameterDescription "Kc_RefPos_y", "Y coordinate of the interaction impedance reference point"
            SetParameterDescription "Kc_RefPos_z", "Z coordinate of the interaction impedance reference point, periodic direction"
    End Select
End Sub

Private Sub ProcessingPhase()

    Dim dPhaseDeg     As Double
    Dim dPhaseRad     As Double
    Dim dBeta         As Double

    If g_iNumModes > 0 And g_nBufModes <> g_iNumModes Then
        ReDim g_aBufBeta(1 To g_iNumModes)
        ReDim g_aBufZpierce(1 To g_iNumModes)
        ReDim g_aBufPhase(1 To g_iNumModes)
        g_nBufModes = g_iNumModes
    End If

    If g_iNumModes > 0 And g_nPhaseStateModes <> g_iNumModes Then
        ReDim g_abPhaseDead(1 To g_iNumModes)
        ReDim g_anPhaseFlat(1 To g_iNumModes)
        ReDim g_abPhaseWarned(1 To g_iNumModes)
        g_nPhaseStateModes = g_iNumModes
    End If

    Dim nParams     As Long
    Dim i           As Integer
    Dim nPhaseIndex As Integer
    Dim bFound      As Boolean
    Dim sParamName  As String
    Dim dParamVal   As Double

    With ParameterSweep
        nParams = .GetNumberOfVaryingParameters

        bFound      = False
        nPhaseIndex = -1
        g_sGroupPath = ""
        g_sGroupKey  = ""

        For i = 0 To nParams - 1
            sParamName = .GetNameOfVaryingParameter(i)

            If sParamName = "phase" Then
                nPhaseIndex = i
                bFound      = True
            Else
                dParamVal = .GetValueOfVaryingParameter(i)

                sParamName = Replace(sParamName, "|", "_")
                sParamName = Replace(sParamName, "\", "_")

                If g_sGroupPath <> "" Then g_sGroupPath = g_sGroupPath & "\"
                g_sGroupPath = g_sGroupPath & sParamName & " = " & Format(dParamVal, "0.0#####")

                If g_sGroupKey <> "" Then g_sGroupKey = g_sGroupKey & " & "
                g_sGroupKey = g_sGroupKey & sParamName & " = " & Format(dParamVal, "0.0#####")
            End If
        Next i

        If Not bFound Then
            LogCritical "Parameter 'phase' not found in the sweep parameter list!", True
            End
        End If

        dPhaseDeg = .GetValueOfVaryingParameter(nPhaseIndex)
    End With

    Dim nGroupIndex As Long, nGroupCount As Long
    nGroupIndex = LookupGroupIndex(g_sGroupKey, nGroupCount)
    If nGroupIndex = 0 Then nGroupIndex = nGroupCount + 1

    Dim fGroup As Integer
    fGroup = FreeFile
    Open g_sTemp & "_groups.txt" For Append As #fGroup
    Print #fGroup, g_sGroupPath & "|" & g_sGroupKey & "|" & nGroupIndex
    Close #fGroup

    g_sFileTag = MakeFileTag(g_sGroupKey, nGroupIndex, g_nFileTagBudget)

    Dim sFileSuffix As String
    If g_sFileTag = "" Then sFileSuffix = "" Else sFileSuffix = "_" & g_sFileTag

    dPhaseRad = dPhaseDeg * DegToRad
    dBeta = dPhaseRad / g_dPitchSI

    If g_sGroupKey <> "" Then LogInfo " ---------- Parameter set: " & g_sGroupKey & " ---------- "

    LogInfo " ----- phase = " & dPhaseDeg & "＜ ----- "

    Dim dPrevFreq() As Double
    ReDim dPrevFreq(1 To g_iNumModes)
    Dim iPrevFile As Integer
    Dim iMode As Integer
    For iMode = 1 To g_iNumModes
        If g_abModeFlag(iMode) Then
            Err.Clear
            On Error Resume Next
            iPrevFile = FreeFile
            Open g_sTemp & "_prevfreq_" & iMode & sFileSuffix & ".txt" For Input As #iPrevFile
            If Err = 0 Then
                Input #iPrevFile, dPrevFreq(iMode)
                If Err <> 0 Then dPrevFreq(iMode) = -1#
                Close #iPrevFile
            Else
                dPrevFreq(iMode) = -1#
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next iMode

    Dim dCurFreq() As Double
    Dim abFreqOK() As Boolean
    ReDim dCurFreq(1 To g_iNumModes)
    ReDim abFreqOK(1 To g_iNumModes)

    Dim bPhaseIsPiMultiple As Boolean
    bPhaseIsPiMultiple = (Abs(Sin(dPhaseDeg * DegToRad)) < 1e-4)

    Dim sMode As String
    Dim dFreqHz As Double, dFreqGHz As Double
    Dim dVre As Double, dVim As Double, dEabs As Double, dVSq As Double
    Dim dPower As Double, dKc As Double

    If g_sGroupKey <> g_sLastGroupKey Then
        g_nCross = 0
        g_bGridFailed = False
        If RowBuffersPending Then
            LogError "Parameter set changed: discarding temporary data rows left unwritten by the previous set"
        End If
        ClearRowBuffers
        ResetPhaseSweepState
        g_sLastGroupKey = g_sGroupKey
    End If

    If Not bPhaseIsPiMultiple Then
        PrecomputeMeshGrid
    Else
        LogInfo "  ***  Phase " & dPhaseDeg & " deg is an integer multiple of Pi  ***"
        LogInfo "  ***  Skipping sampling-grid / point-file generation and field integration / power flow / interaction impedance  ***"
    End If

    Mesh.ViewMeshMode False

    For iMode = 1 To g_iNumModes
        dFreqHz = 0#

        sMode = CStr(iMode)

        Dim bModeSelected As Boolean
        bModeSelected = g_abModeFlag(iMode)
        If Not bModeSelected Then
            LogInfo "  Mode " & sMode & " not selected, skipped"
            GoTo SkipModeCalc
        End If

        Dim iRetry As Integer
        Dim bFreqOK As Boolean
        iRetry = 0
        bFreqOK = False

        dFreqHz = R3DGetFieldFrequency(sMode)
        If dFreqHz > 0# Then bFreqOK = True

        Do While Not bFreqOK And iRetry < MAX_FREQ_RETRY
            iRetry = iRetry + 1
            SwitchToItem "2D/3D Results\Modes\Mode " & sMode & "\e"
            dFreqHz = GetFieldFrequency()
            If dFreqHz > 0# Then
                bFreqOK = True
                Exit Do
            End If
            DoEvents
            Wait FREQ_RETRY_WAIT
        Loop

        If Not bFreqOK Then
            LogError "  Phase " & dPhaseDeg & " deg, Mode " & sMode & _
                ": failed to read the field frequency after " & iRetry & " retries; skipping this point"
            GoTo SkipModeCalc
        End If

        dFreqGHz = dFreqHz * Units.GetFrequencyUnitToSI / 1e9
        dCurFreq(iMode) = dFreqHz
        abFreqOK(iMode) = True

        Dim bFreqDirectionOK As Boolean
        Dim bFreqFlat      As Boolean
        Dim dFreqRel       As Double

        bFreqDirectionOK = True
        bFreqFlat = False
        dFreqRel = 1#

        If dPrevFreq(iMode) >= 0# Then
            Select Case g_iRegionMode
                Case REGION_FORWARD
                    bFreqDirectionOK = (dFreqHz > dPrevFreq(iMode))
                Case REGION_BACKWARD
                    bFreqDirectionOK = (dFreqHz < dPrevFreq(iMode))
            End Select

            If dPrevFreq(iMode) > 0# Then
                dFreqRel = Abs(dFreqHz - dPrevFreq(iMode)) / dPrevFreq(iMode)
                bFreqFlat = (dFreqRel <= FREQ_FLAT_TOL)
            End If
        End If

        If bFreqFlat Then
            g_anPhaseFlat(iMode) = g_anPhaseFlat(iMode) + 1
            If Not g_abPhaseWarned(iMode) Then
                g_abPhaseWarned(iMode) = True
                LogWarning "  Mode " & sMode & ": frequency does not change with phase (f = " & _
                    Format(dFreqHz, "0.000E+00") & " " & Units.GetFrequencyUnit & ", relative change " & _
                    Format(dFreqRel, "0.0E+00") & "). Usual cause: the solver is not applying the phase difference to the " & _
                    g_sDirLabel & " direction periodic boundary (every sweep point solves the same mode), or the mode" & _
                    " is close to a Pi point. In both cases the net power flow is near zero, so the interaction impedance" & _
                    " is meaningless and has been marked invalid"
            ElseIf g_anPhaseFlat(iMode) <= FREQ_FLAT_MAX Then
                LogInfo "  Mode " & sMode & ": frequency still does not change with phase (" & _
                    g_anPhaseFlat(iMode) & " consecutive points); the interaction impedance at this point is marked invalid"
            End If
            If g_anPhaseFlat(iMode) >= FREQ_FLAT_MAX And Not g_abPhaseDead(iMode) Then
                g_abPhaseDead(iMode) = True
                LogWarning "  Mode " & sMode & ": the frequency is identical for " & FREQ_FLAT_MAX & " consecutive sweep points, " & _
                    "so this phase sweep is judged invalid. Check the " & g_sDirLabel & " direction: " & _
                    "(1) whether both faces are set to periodic; (2) whether the periodic phase difference is bound to the sweep variable 'phase'; " & _
                    "(3) whether the computation domain spans exactly one period along that direction."
                LogWarning "  Mode " & sMode & ": subsequent sweep points skip field export and integration; only frequency and phase are recorded"
            End If
        Else
            g_anPhaseFlat(iMode) = 0
            If g_abPhaseDead(iMode) Then
                g_abPhaseDead(iMode) = False
                g_abPhaseWarned(iMode) = False
                LogInfo "  Mode " & sMode & ": frequency changes with phase again; interaction impedance calculation is restored for this mode"
            End If
        End If

        Dim bSkipImpedance As Boolean
        bSkipImpedance = bPhaseIsPiMultiple Or (Not bFreqDirectionOK) Or g_abPhaseDead(iMode)

        If bSkipImpedance Then
            dVre = 0#: dVim = 0#: dEabs = 0#: dVSq = 0#
            dPower = -1#
            dKc = -1#

            LogInfo "  Mode " & sMode & ": freq = " & Format(dFreqGHz, "0.000000") & " GHz"

            If Not bPhaseIsPiMultiple And Not bFreqFlat Then
                If g_abPhaseDead(iMode) Then
                    LogInfo "  Mode " & sMode & ": phase sweep judged invalid; skipping field export / field integration / power flow / interaction impedance"
                ElseIf Not bFreqDirectionOK Then
                    Select Case g_iRegionMode
                        Case REGION_FORWARD
                            LogInfo "  *** Frequency did not rise: " & Format(dPrevFreq(iMode), "0.000E+00") & _
                                " ★ " & Format(dFreqHz, "0.000E+00") & " " & Units.GetFrequencyUnit & _
                                ", this point is not in the forward-wave region; skipping field integration / power flow / interaction impedance  ***"
                        Case REGION_BACKWARD
                            LogInfo "  *** Frequency did not fall: " & Format(dPrevFreq(iMode), "0.000E+00") & _
                                " ★ " & Format(dFreqHz, "0.000E+00") & " " & Units.GetFrequencyUnit & _
                                ", this point is not in the backward-wave region; skipping field integration / power flow / interaction impedance  ***"
                    End Select
                End If
            End If
        Else
            LogInfo "  Mode " & sMode & ": freq = " & Format(dFreqGHz, "0.000000") & _
                " GHz, computing..."

            Dim bFieldsOK As Boolean
            If g_iPowerFlowSrc = 1 Then
                bFieldsOK = PreparePowerFieldFromCST(iMode, sMode)
            Else
                bFieldsOK = LoadFieldDataFromEH(iMode)
            End If

            If bFieldsOK Then
                IntegrateLongitudinal dBeta, dVre, dVim, dEabs
                dVSq = dVre * dVre + dVim * dVim
                LogInfo "    Longitudinal integral: V_avg = (" & Format(dVre, "0.000E+00") & ", " & _
                    Format(dVim, "0.000E+00") & "), |V|^2 = " & Format(dVSq, "0.000E+00") & _
                    ", E_abs = " & Format(dEabs, "0.000E+00")

                If g_iPowerFlowSrc = 1 Then
                    ComputePowerFlowFromCST dPower
                Else
                    ComputePowerFlowFromEH dPower
                End If
                LogInfo "    Power flow: P = " & Format(dPower, "0.000E+00") & " W"

                If g_iPowerFlowSrc = 0 And dPower = 0# Then
                    LogWarning "    Note: the power flow is 0 with the E/H method. Make sure the H field of Mode " & sMode & _
                        " can be selected in the result tree -- if the H-field export fails, Re(E x H*) stays 0."
                End If

                If dVSq = 0# Then
                    dKc = -1#
                    LogWarning "    No valid longitudinal integral (|V|^2 = 0); the interaction impedance at this point is marked invalid"
                ElseIf dBeta = 0# Or dPower <= 0# Then
                    dKc = -1#
                    If dBeta = 0# Then
                        LogWarning "    Phase constant is zero (beta = 0); the interaction impedance at this point is marked invalid"
                    Else
                        LogWarning "    Power flow is not positive (P = " & Format(dPower, "0.000E+00") & " W); the interaction impedance at this point is marked invalid"
                    End If
                Else
                    dKc = dVSq / (2# * dBeta * dBeta * dPower)
                    If dKc > KC_SANITY_MAX Then
                        If Not g_abPhaseWarned(iMode) Then
                            g_abPhaseWarned(iMode) = True
                            LogWarning "    Power flow is abnormally small: P = " & Format(dPower, "0.000E+00") & " W, giving Kc = " & _
                                Format(dKc, "0.000E+00") & " Ohm, above the sanity limit " & _
                                Format(KC_SANITY_MAX, "0.000E+00") & " Ohm; the interaction impedance at this point is marked invalid"
                        Else
                            LogInfo "    Kc exceeds the sanity limit (" & Format(dKc, "0.000E+00") & " Ohm); this point is marked invalid"
                        End If
                        dKc = -1#
                    End If
                End If
                LogInfo "    Interaction impedance: Kc = " & Format(dKc, "0.000E+00") & " Ohm"
            Else
                LogError "Mode " & sMode & ": field export / parsing failed; invalid values are written for this point"
                dVre = 0#: dVim = 0#: dEabs = 0#: dVSq = 0#
                dPower = -1#: dKc = -1#
            End If
        End If

        BufferAppend g_aBufBeta(iMode),    dFreqGHz, dBeta
        BufferAppend g_aBufZpierce(iMode), dFreqGHz, dKc
        BufferAppend g_aBufPhase(iMode),   dFreqGHz, dPhaseDeg
SkipModeCalc:
    Next iMode

    For iMode = 1 To g_iNumModes
        bModeSelected = g_abModeFlag(iMode)
        If bModeSelected Then
            BufferFlush TempPath(F_BETA    & sFileSuffix, iMode), g_aBufBeta(iMode)
            BufferFlush TempPath(F_ZPIERCE & sFileSuffix, iMode), g_aBufZpierce(iMode)
            BufferFlush TempPath(F_PHASE   & sFileSuffix, iMode), g_aBufPhase(iMode)
        End If
    Next iMode

    For iMode = 1 To g_iNumModes
        If g_abModeFlag(iMode) And abFreqOK(iMode) Then
            iPrevFile = FreeFile
            Open g_sTemp & "_prevfreq_" & iMode & sFileSuffix & ".txt" For Output As #iPrevFile
            Print #iPrevFile, dCurFreq(iMode)
            Close #iPrevFile
        End If
    Next iMode

End Sub

Private Sub FinalizationPhase()
    LogInfo "======================================================================="
    LogInfo "  Sweep finished, collecting results..."
    LogInfo "======================================================================="

    Dim iMode As Integer, sMode As String
    Dim oBeta As Object, oZpierce As Object, oVphase As Object
    Dim oBetaPhase As Object

    Dim g As Integer
    Dim sPath As String
    Dim sKey  As String
    Dim sFileSuffix2 As String
    Dim sTreeBase As String

    Dim aPath() As String, aKey() As String
    Dim aIdx() As Long
    Dim nG As Long
    nG = 0
    Dim sLine As String, sP As String, sK As String, sIdx As String
    Dim nPos As Long, nPos2 As Long, k As Long, bDup As Boolean
    Dim fGroup As Integer

    On Error Resume Next
    If Dir(g_sTemp & "_groups.txt") <> "" Then
        fGroup = FreeFile
        Open g_sTemp & "_groups.txt" For Input As #fGroup
        Do While Not EOF(fGroup)
            Line Input #fGroup, sLine
            If Trim(sLine) <> "" Then
                nPos = InStr(sLine, "|")
                If nPos > 0 Then
                    sP = Left(sLine, nPos - 1)
                    sK = Mid(sLine, nPos + 1)
                    nPos2 = InStr(sK, "|")
                    If nPos2 > 0 Then
                        sIdx = Mid(sK, nPos2 + 1)
                        sK = Left(sK, nPos2 - 1)
                    Else
                        sIdx = ""
                    End If
                Else
                    sP = sLine
                    sK = sLine
                    sIdx = ""
                End If

                bDup = False
                For k = 1 To nG
                    If aPath(k) = sP Then bDup = True: Exit For
                Next k
                If Not bDup Then
                    nG = nG + 1
                    ReDim Preserve aPath(1 To nG)
                    ReDim Preserve aKey(1 To nG)
                    ReDim Preserve aIdx(1 To nG)
                    aPath(nG) = sP
                    aKey(nG) = sK
                    aIdx(nG) = CLng(Val(sIdx))
                    If aIdx(nG) = 0 Then aIdx(nG) = nG
                End If
            End If
        Loop
        Close #fGroup
    End If
    On Error GoTo 0

    For g = 1 To nG
        sPath = aPath(g)
        sKey  = aKey(g)

        sFileSuffix2 = MakeFileTag(sKey, aIdx(g), g_nFileTagBudget)
        If sFileSuffix2 <> "" Then sFileSuffix2 = "_" & sFileSuffix2

        If sPath = "" Then
            sTreeBase = "1D Results\" & F_RESULT_GROUP
        Else
            sTreeBase = "1D Results\" & F_RESULT_GROUP & "\" & sPath
        End If

        For iMode = 1 To g_iNumModes
            sMode = CStr(iMode)
            Dim bModeSelected As Boolean
            bModeSelected = g_abModeFlag(iMode)

            If Not bModeSelected Then
                LogInfo "  Mode " & sMode & ": not selected, skipped"
                GoTo SkipModeFinalize
            End If

            If sPath = "" Then
                LogInfo "---------- Mode " & sMode & ": collecting result curves ----------"
            Else
                LogInfo "---------- Parameter set [ " & Replace(sPath, "\", " & ") & " ] Mode " & sMode & ": collecting result curves ----------"
            End If

            Dim sCurBeta As String, sCurPhase As String, sCurVp As String, sCurKc As String
            sCurBeta  = g_sTemp & "_curveBeta"  & sFileSuffix2 & "_" & sMode & ".txt"
            sCurPhase = g_sTemp & "_curvePhase" & sFileSuffix2 & "_" & sMode & ".txt"
            sCurVp    = g_sTemp & "_curveVp"    & sFileSuffix2 & "_" & sMode & ".txt"
            sCurKc    = g_sTemp & "_curveKc"    & sFileSuffix2 & "_" & sMode & ".txt"

            Dim nBetaPt As Long, nPhasePt As Long, nVpPt As Long, nKcPt As Long
            Dim nRetCurve As Long
            nRetCurve = CBuildCurveFiles(CORE_AUTHOR, g_sTemp, sFileSuffix2, iMode, CLight, _
                sCurBeta, sCurPhase, sCurVp, sCurKc, _
                nBetaPt, nPhasePt, nVpPt, nKcPt)

            If nRetCurve <> 0 Then
                LogError "Mode " & sMode & ": no valid sweep data (curve generation failed, return code " & nRetCurve & "); skipping this mode"
                GoTo SkipModeFinalize
            End If

            If nKcPt < nBetaPt Then
                LogInfo "  Mode " & sMode & ": " & (nBetaPt - nKcPt) & _
                    " sweep points have no valid interaction impedance (no |V|^2 result, or the frequency direction does not match the selected region); skipped"
            End If

            If nBetaPt > 0 Then
                Set oBeta = Result1D("")
                oBeta.LoadPlainFile sCurBeta
                Save1D oBeta, "Brillouin Diagram", _
                    g_sResult & F_BETA & "_" & sMode & sFileSuffix2 & ".sig", _
                    sTreeBase & "\Brillouin Diagram Beta\Mode " & sMode, _
                    "Beta (rad/m)", "Frequency (GHz)"
                LogInfo "  Mode " & sMode & ": Brillouin (Beta) curve saved"
            End If

            If nPhasePt > 0 Then
                Set oBetaPhase = Result1D("")
                oBetaPhase.LoadPlainFile sCurPhase
                Save1D oBetaPhase, "Brillouin Diagram (Phase)", _
                    g_sResult & F_PHASE & "_" & sMode & sFileSuffix2 & ".sig", _
                    sTreeBase & "\Brillouin Diagram Phase\Mode " & sMode, _
                    "Phase (deg)", "Frequency (GHz)"
                LogInfo "  Mode " & sMode & ": Brillouin (Phase) curve saved"
            End If

            If nVpPt > 0 Then
                Set oVphase = Result1D("")
                oVphase.LoadPlainFile sCurVp
                Save1D oVphase, "Normalized Phase Velocity Plot", _
                    g_sResult & F_VPHASE & "_" & sMode & sFileSuffix2 & ".sig", _
                    sTreeBase & "\Normalized Phase Velocity\Mode " & sMode, _
                    , "Normalized Phase Velocity (clight)"
            End If

            If nKcPt > 0 Then
                Set oZpierce = Result1D("")
                oZpierce.LoadPlainFile sCurKc
                Save1D oZpierce, "Pierce Interaction Impedance Plot", _
                    g_sResult & F_ZPIERCE & "_" & sMode & sFileSuffix2 & ".sig", _
                    sTreeBase & "\Pierce Interaction Impedance\Mode " & sMode, _
                    , "Pierce Interaction Impedance (Ohm)"
            End If

            Set oBeta      = Nothing
            Set oZpierce   = Nothing
            Set oVphase    = Nothing
            Set oBetaPhase = Nothing

SkipModeFinalize:
        Next iMode
    Next g

    LogInfo "======================================================================="

    For g = 1 To nG
        sKey = aKey(g)
        sFileSuffix2 = MakeFileTag(sKey, aIdx(g), g_nFileTagBudget)
        If sFileSuffix2 <> "" Then sFileSuffix2 = "_" & sFileSuffix2

        For iMode = 1 To g_iNumModes
            On Error Resume Next
            Kill TempPath(F_BETA    & sFileSuffix2, iMode)
            Kill TempPath(F_ZPIERCE & sFileSuffix2, iMode)
            Kill TempPath(F_PHASE   & sFileSuffix2, iMode)
            On Error GoTo 0
        Next iMode
    Next g

    For iMode = 1 To g_iNumModes
        On Error Resume Next
        Kill g_sTemp & "_E_mode" & CStr(iMode) & ".txt"
        Kill g_sTemp & "_H_mode" & CStr(iMode) & ".txt"
        Kill g_sTemp & "_P_mode" & CStr(iMode) & ".txt"
        On Error GoTo 0
        LogInfo "  Mode " & iMode & ": temporary files cleaned up"
    Next iMode

    On Error Resume Next
    Kill g_sPointsFile
    Kill g_sTemp & "_groups.txt"
    DeleteFilesByPattern g_sTemp & "_prevfreq_*.txt"
    DeleteFilesByPattern g_sTemp & "_curve*.txt"
    SelectTreeItem "1D Results\" & F_RESULT_GROUP
    On Error GoTo 0

    Dim sStart As String
    Dim dElapsed As Double
    Dim bHaveStart As Boolean
    bHaveStart = False
    Dim fTime2 As Integer
    fTime2 = FreeFile
    Err.Clear
    On Error Resume Next
    Open g_sTemp & "_runtime_start.txt" For Input As #fTime2
    If Err = 0 Then
        Line Input #fTime2, sStart
        If Err = 0 Then bHaveStart = True
        Close #fTime2
    End If
    Err.Clear
    On Error GoTo 0

    If bHaveStart Then
        dElapsed = DateDiff("s", CDate(sStart), Now)
        On Error Resume Next
        Kill g_sTemp & "_runtime_start.txt"
        On Error GoTo 0
    End If

    Dim nTotalSec As Long, nDay As Long, nHour As Long, nMin As Long, nSec As Long
    Dim sTime As String
    If bHaveStart Then
        nTotalSec = CLng(dElapsed)

        nDay  = nTotalSec \ 86400
        nHour = (nTotalSec Mod 86400) \ 3600
        nMin  = (nTotalSec Mod 3600) \ 60
        nSec  = nTotalSec Mod 60

        sTime = ""
        If nDay > 0 Then sTime = sTime & nDay & " d "
        If nDay > 0 Or nHour > 0 Then sTime = sTime & nHour & " h "
        If nDay > 0 Or nHour > 0 Or nMin > 0 Then sTime = sTime & nMin & " m "
        sTime = sTime & nSec & " s"
    Else
        sTime = "(start timestamp missing, elapsed time unavailable)"
    End If

    LogInfo "======================================================================="
    LogInfo "  All done! Processed " & nG & " parameter sets x " & g_iNumModes & " modes"
    LogInfo "  Interaction impedance reference point: X = " & g_xRef & ", Y = " & g_yRef & _
        ", Z = " & g_zRef & ",  units: " & g_sUnit & ", periodic direction: " & g_sDirLabel
    LogInfo "  Power-flow method: " & PowerFlowSrcText(g_iPowerFlowSrc) & _
        ",  region: " & RegionText(g_iRegionMode)
    LogInfo "  Log file: " & g_sLogFile
    If Len(g_sLogLastError) > 0 Then
        LogWarning "  A log write failure occurred during this run: " & g_sLogLastError
    End If
    LogInfo "  Elapsed time: " & sTime
    LogInfo "  End time: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LogInfo "======================================================================="

End Sub

Private Sub LoadMacroConfig()
    Dim sMiss As String

    g_sCfgVerRaw = ""
    g_sCfgLblRaw = ""
    g_sCfgDllMinRaw = ""
    g_nDllMinVer = 0
    g_sAuthorName = ""
    g_sReleaseDate = ""
    g_sPlatform = ""
    g_sCstVerMin = ""
    g_sCstVerMax = ""

    g_bCfgOK = False
    g_sCfgNote = ""
    g_sCfgPath = FindConfigFile()
    sMiss = ""

    If g_sCfgPath = "" Then
        g_sCfgNote = "the external configuration file was not found in any of the 4 candidate locations: " & CFG_FILE_NAME
    ElseIf Not ReadConfigFile(g_sCfgPath) Then
        g_sCfgNote = "the external configuration file cannot be read (in use or insufficient permissions): " & g_sCfgPath
    Else
        sMiss = CfgMissingKeys()
        If Len(sMiss) > 0 Then
            g_sCfgNote = "the configuration file is missing required [Version] keys: " & sMiss & " --- " & g_sCfgPath
        ElseIf Not ParseDllMinVersion() Then
            g_sCfgNote = "DllMinVersion in the configuration file is not a valid version number: """ & _
                g_sCfgDllMinRaw & """ (a positive integer such as 122 is expected) --- " & g_sCfgPath
        Else
            g_sVersionString = g_sCfgVerRaw
            g_sVersionLabel = g_sCfgLblRaw
            g_bCfgOK = True
        End If
    End If
End Sub

Private Function ParseDllMinVersion() As Boolean
    Dim n As Long

    ParseDllMinVersion = False
    g_nDllMinVer = 0

    If Not IsNumeric(g_sCfgDllMinRaw) Then Exit Function

    n = CLng(Val(g_sCfgDllMinRaw))
    If n <= 0 Then Exit Function

    g_nDllMinVer = n
    ParseDllMinVersion = True
End Function

Private Function CfgMissingKeys() As String
    Dim sMiss As String

    sMiss = ""
    If Len(g_sCfgVerRaw) = 0 Then sMiss = sMiss & "VersionString "
    If Len(g_sCfgLblRaw) = 0 Then sMiss = sMiss & "VersionLabel "
    If Len(g_sAuthorName) = 0 Then sMiss = sMiss & "Author "
    If Len(g_sReleaseDate) = 0 Then sMiss = sMiss & "ReleaseDate "
    If Len(g_sPlatform) = 0 Then sMiss = sMiss & "Platform "
    If Len(g_sCstVerMin) = 0 Then sMiss = sMiss & "CstVersionMin "
    If Len(g_sCstVerMax) = 0 Then sMiss = sMiss & "CstVersionMax "
    If Len(g_sCfgDllMinRaw) = 0 Then sMiss = sMiss & "DllMinVersion "

    CfgMissingKeys = Trim$(sMiss)
End Function

Private Function CfgCandidate(ByVal iIndex As Integer) As String
    Dim sPath As String

    sPath = ""
    Err.Clear
    On Error Resume Next
    Select Case iIndex
        Case 1
            sPath = GetInstallPath & "\AMD64\" & CFG_FILE_NAME
        Case 2
            sPath = GetInstallPath & "\Library\Macros\Solver\E-Solver\" & CFG_FILE_NAME
        Case 3
            sPath = GetProjectPath("Root") & "\" & CFG_FILE_NAME
        Case 4
            sPath = GetProjectPath("Project") & "\" & CFG_FILE_NAME
    End Select
    If Err.Number <> 0 Then
        sPath = ""
        Err.Clear
    End If
    On Error GoTo 0

    CfgCandidate = sPath
End Function

Private Function CfgFileExists(ByVal sPath As String) As Boolean
    Dim sHit As String

    CfgFileExists = False
    sHit = ""
    Err.Clear
    On Error Resume Next
    sHit = Dir(sPath)
    If Err.Number = 0 And Len(sHit) > 0 Then CfgFileExists = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function FindConfigFile() As String
    Dim i As Integer
    Dim sPath As String

    FindConfigFile = ""
    For i = 1 To 4
        sPath = CfgCandidate(i)
        If Len(sPath) > 0 Then
            If CfgFileExists(sPath) Then
                FindConfigFile = sPath
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ReadConfigFile(ByVal sPath As String) As Boolean
    Dim f As Integer
    Dim sLine As String
    Dim sSection As String
    Dim sKey As String
    Dim sVal As String
    Dim nPos As Long
    Dim bOpened As Boolean

    ReadConfigFile = False
    sSection = ""
    bOpened = False

    Err.Clear
    On Error Resume Next
    f = FreeFile
    Open sPath For Input As #f
    If Err.Number = 0 Then bOpened = True
    Err.Clear
    On Error GoTo 0
    If Not bOpened Then Exit Function

    Err.Clear
    On Error Resume Next
    Do While Not EOF(f)
        Line Input #f, sLine
        If Err.Number <> 0 Then Exit Do
        sLine = Trim$(sLine)
        If Len(sLine) > 0 Then
            If Left$(sLine, 1) = ";" Or Left$(sLine, 1) = "#" Then
            ElseIf Left$(sLine, 1) = "[" Then
                nPos = InStr(sLine, "]")
                If nPos > 1 Then sSection = UCase$(Trim$(Mid$(sLine, 2, nPos - 2)))
            Else
                nPos = InStr(sLine, "=")
                If nPos > 0 Then
                    sKey = UCase$(Trim$(Left$(sLine, nPos - 1)))
                    sVal = Trim$(Mid$(sLine, nPos + 1))
                    ApplyConfigKey sSection, sKey, sVal
                End If
            End If
        End If
    Loop
    Err.Clear
    Close #f
    On Error GoTo 0

    ReadConfigFile = True
End Function

Private Sub ApplyConfigKey(ByVal sSection As String, ByVal sKey As String, ByVal sVal As String)
    If Len(sVal) = 0 Then Exit Sub
    If sSection <> CFG_SECTION_VER Then Exit Sub

    Select Case sKey
        Case "VERSIONSTRING"
            g_sCfgVerRaw = sVal
        Case "VERSIONLABEL"
            g_sCfgLblRaw = sVal
        Case "AUTHOR"
            g_sAuthorName = sVal
        Case "RELEASEDATE"
            g_sReleaseDate = sVal
        Case "PLATFORM"
            g_sPlatform = sVal
        Case "CSTVERSIONMIN"
            g_sCstVerMin = sVal
        Case "CSTVERSIONMAX"
            g_sCstVerMax = sVal
        Case "DLLMINVERSION"
            g_sCfgDllMinRaw = sVal
    End Select
End Sub

Private Sub AbortOnConfigError()
    Dim sMsg As String
    Dim i As Integer

    sMsg = "Error: the macro version information must come from the external configuration file " & CFG_FILE_NAME & _
        ", which is currently unavailable, so the macro has been aborted!" & vbCrLf & vbCrLf & _
        "Reason: " & g_sCfgNote & vbCrLf & vbCrLf & _
        "The macro searches for the file in the following order; make sure at least one copy exists and is readable and that the [Version] section has all 8" & vbCrLf & _
        " keys (no fallback defaults remain in the code):" & vbCrLf
    For i = 1 To 4
        If Len(CfgCandidate(i)) > 0 Then sMsg = sMsg & "  " & i & ". " & CfgCandidate(i) & vbCrLf
    Next i
    sMsg = sMsg & vbCrLf & _
        "Note: this file must be saved as ANSI/GBK; saving it as UTF-8 makes every key unreadable." & vbCrLf & _
        "      Re-running the install script (Double-click me for automatic installation.bat) restores the default deployment."

    MsgBox sMsg, vbCritical, "Macro aborted"
    On Error Resume Next
    LogCritical sMsg, True
    On Error GoTo 0
    End
End Sub

Private Sub CheckCoreLibrary()
    Dim sDllPath As String
    sDllPath = GetInstallPath & "\AMD64\LinXi.dll"

    If Dir(sDllPath) = "" Then
        MsgBox "Error: cannot load the core computation library LinXi.dll." & vbCrLf & _
            "Core computation library file not found:" & vbCrLf & sDllPath & vbCrLf & vbCrLf & _
            "Please check:" & vbCrLf & _
            "  1. that the core computation library LinXi.dll is deployed correctly;" & vbCrLf & _
            "  2. that CST is a 64-bit build (the DLL is 64-bit).", vbCritical
        LogCritical "Failed to load the core computation library LinXi.dll; the file does not exist: " & sDllPath, True
        End
    End If

    Dim nVer As Long
    Dim nErrNum As Long, nDllErr As Long
    Dim sDesc As String
    Err.Clear
    On Error Resume Next
    nVer = CCoreVersion(CORE_AUTHOR)
    nErrNum = Err.Number
    nDllErr = Err.LastDLLError
    sDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    If nVer < g_nDllMinVer Then
        MsgBox "Error: failed to load the core computation library." & vbCrLf & vbCrLf & _
            "The file exists but cannot be called correctly. Possible causes:" & vbCrLf & _
            "  1. the DLL is 32-bit (it must be 64-bit);" & vbCrLf & _
            "  2. the exported function names do not match;" & vbCrLf & _
            "  3. a library the DLL depends on is missing; check or update the DLL version." & vbCrLf & vbCrLf & _
            "Reported version: v" & nVer & ", this program requires at least v" & g_nDllMinVer & vbCrLf & _
            "(the minimum version requirement is written in " & CFG_FILE_NAME & " DllMinVersion)" & vbCrLf & vbCrLf & _
            "Err.Number = " & nErrNum & vbCrLf & _
            "LastDLLError = " & nDllErr & vbCrLf & _
            "Description = " & sDesc, vbCritical
        LogCritical " -----  Failed to load the core computation library: Err=" & nErrNum & ", LastDLLError=" & nDllErr & _
            ", reported version=" & nVer & ", required at least=" & g_nDllMinVer & "  ----- ", True
        End
    End If

    LogInfo " ----- Core computation library LinXi.dll loaded successfully, version v" & nVer & " ----- "

    If g_bCfgOK Then
        LogInfo " ----- CST Eigenmode Slow-Wave Structure User Watch Macro " & g_sVersionString & " (" & g_sVersionLabel & ")" &" ----- "
    Else
        LogWarning " ----- Version information could not be read from " & CFG_FILE_NAME & ": " & g_sCfgNote & " ----- "
    End If

End Sub

Private Function DialogFunc(ByVal DlgItem$, ByVal Action%, ByVal SuppValue&) As Boolean
    Select Case Action
        Case 1
            DlgEnable "selectMode1",  0
            DlgEnable "selectMode2",  0
            DlgEnable "selectMode3",  0
            DlgEnable "selectMode4",  0
            DlgEnable "selectMode5",  0
            DlgEnable "selectMode6",  0
            DlgEnable "selectMode7",  0
            DlgEnable "selectMode8",  0
            DlgEnable "selectMode9",  0
            DlgEnable "selectMode10", 0
            DlgEnable "selectMode11", 0
            DlgEnable "selectMode12", 0
            DlgEnable "selectMode13", 0
            DlgEnable "selectMode14", 0
            DlgEnable "selectMode15", 0
            DlgEnable "selectMode16", 0
            DlgEnable "selectMode17", 0
            DlgEnable "selectMode18", 0
            DlgEnable "selectMode19", 0
            DlgEnable "selectMode20", 0
            DlgVisible "txtMode",     0
            DlgVisible "lblHint",     0

            Select Case g_sDirLabel
                Case "X"
                    DlgEnable "KcPosx", 0
                    DlgText   "KcPosx", "domain center"
                Case "Y"
                    DlgEnable "KcPosy", 0
                    DlgText   "KcPosy", "domain center"
                Case "Z"
                    DlgEnable "KcPosz", 0
                    DlgText   "KcPosz", "domain center"
            End Select

            Select Case DlgValue("Group1")
                Case 0
                Case 1
                    DlgEnable "selectMode1",  1
                    DlgEnable "selectMode2",  1
                    DlgEnable "selectMode3",  1
                    DlgEnable "selectMode4",  1
                    DlgEnable "selectMode5",  1
                    DlgEnable "selectMode6",  1
                    DlgEnable "selectMode7",  1
                    DlgEnable "selectMode8",  1
                    DlgEnable "selectMode9",  1
                    DlgEnable "selectMode10", 1
                    DlgEnable "selectMode11", 1
                    DlgEnable "selectMode12", 1
                    DlgEnable "selectMode13", 1
                    DlgEnable "selectMode14", 1
                    DlgEnable "selectMode15", 1
                    DlgEnable "selectMode16", 1
                    DlgEnable "selectMode17", 1
                    DlgEnable "selectMode18", 1
                    DlgEnable "selectMode19", 1
                    DlgEnable "selectMode20", 1
                    DlgVisible "lblHint",     1
                    DlgText   "lblHint", "Note: tick the modes whose interaction impedance is needed"
                Case 2
                    DlgVisible "txtMode",     1
                    DlgVisible "lblHint",     1
                    DlgText   "lblHint", "Note: separate the numbers with commas, e.g. 1,3,5,7"
                    DlgEnable "txtMode", 1
                    DlgFocus  "txtMode"
            End Select

            If DlgValue("Group2") = 1 Then
                DlgVisible "lblPwDesc1", 0
                DlgVisible "lblPwDesc2", 1
            Else
                DlgVisible "lblPwDesc1", 1
                DlgVisible "lblPwDesc2", 0
            End If

            Select Case DlgValue("Group3")
                Case REGION_FORWARD
                    DlgVisible "lblRgnDesc0", 0
                    DlgVisible "lblRgnDesc1", 1
                    DlgVisible "lblRgnDesc2", 0
                Case REGION_BACKWARD
                    DlgVisible "lblRgnDesc0", 0
                    DlgVisible "lblRgnDesc1", 0
                    DlgVisible "lblRgnDesc2", 1
                Case Else
                    DlgVisible "lblRgnDesc0", 1
                    DlgVisible "lblRgnDesc1", 0
                    DlgVisible "lblRgnDesc2", 0
            End Select

            DialogFunc = False

        Case 2
            Select Case DlgItem$
                Case "Group1"
                    Select Case SuppValue
                        Case 0
                            DlgEnable "selectMode1",  0
                            DlgEnable "selectMode2",  0
                            DlgEnable "selectMode3",  0
                            DlgEnable "selectMode4",  0
                            DlgEnable "selectMode5",  0
                            DlgEnable "selectMode6",  0
                            DlgEnable "selectMode7",  0
                            DlgEnable "selectMode8",  0
                            DlgEnable "selectMode9",  0
                            DlgEnable "selectMode10", 0
                            DlgEnable "selectMode11", 0
                            DlgEnable "selectMode12", 0
                            DlgEnable "selectMode13", 0
                            DlgEnable "selectMode14", 0
                            DlgEnable "selectMode15", 0
                            DlgEnable "selectMode16", 0
                            DlgEnable "selectMode17", 0
                            DlgEnable "selectMode18", 0
                            DlgEnable "selectMode19", 0
                            DlgEnable "selectMode20", 0
                            DlgVisible "txtMode",     0
                            DlgVisible "lblHint",     0
                            DlgText   "txtMode", ""
                        Case 1
                            DlgEnable "selectMode1",  1
                            DlgEnable "selectMode2",  1
                            DlgEnable "selectMode3",  1
                            DlgEnable "selectMode4",  1
                            DlgEnable "selectMode5",  1
                            DlgEnable "selectMode6",  1
                            DlgEnable "selectMode7",  1
                            DlgEnable "selectMode8",  1
                            DlgEnable "selectMode9",  1
                            DlgEnable "selectMode10", 1
                            DlgEnable "selectMode11", 1
                            DlgEnable "selectMode12", 1
                            DlgEnable "selectMode13", 1
                            DlgEnable "selectMode14", 1
                            DlgEnable "selectMode15", 1
                            DlgEnable "selectMode16", 1
                            DlgEnable "selectMode17", 1
                            DlgEnable "selectMode18", 1
                            DlgEnable "selectMode19", 1
                            DlgEnable "selectMode20", 1
                            DlgVisible "txtMode",     0
                            DlgVisible "lblHint",     1
                            DlgText   "lblHint", "Note: tick the modes whose interaction impedance is needed"
                            DlgFocus  "selectMode1"
                        Case 2
                            DlgEnable "selectMode1",  0
                            DlgEnable "selectMode2",  0
                            DlgEnable "selectMode3",  0
                            DlgEnable "selectMode4",  0
                            DlgEnable "selectMode5",  0
                            DlgEnable "selectMode6",  0
                            DlgEnable "selectMode7",  0
                            DlgEnable "selectMode8",  0
                            DlgEnable "selectMode9",  0
                            DlgEnable "selectMode10", 0
                            DlgEnable "selectMode11", 0
                            DlgEnable "selectMode12", 0
                            DlgEnable "selectMode13", 0
                            DlgEnable "selectMode14", 0
                            DlgEnable "selectMode15", 0
                            DlgEnable "selectMode16", 0
                            DlgEnable "selectMode17", 0
                            DlgEnable "selectMode18", 0
                            DlgEnable "selectMode19", 0
                            DlgEnable "selectMode20", 0
                            DlgVisible "txtMode",     1
                            DlgVisible "lblHint",     1
                            DlgText   "lblHint", "Note: separate the numbers with commas, e.g. 1,3,5,7"
                            DlgText   "txtMode", ""
                            DlgEnable "txtMode", 1
                            DlgFocus  "txtMode"
                    End Select
                    DialogFunc = True
                Case "selectMode1", "selectMode2", "selectMode3", "selectMode4", _
                     "selectMode5", "selectMode6", "selectMode7", "selectMode8", _
                     "selectMode9", "selectMode10", "selectMode11", "selectMode12", _
                     "selectMode13", "selectMode14", "selectMode15", "selectMode16", _
                     "selectMode17", "selectMode18", "selectMode19", "selectMode20"
                    DlgValue "Group1", 1
                    DialogFunc = True
                Case "Group2"
                    If SuppValue = 1 Then
                        DlgVisible "lblPwDesc1", 0
                        DlgVisible "lblPwDesc2", 1
                    Else
                        DlgVisible "lblPwDesc1", 1
                        DlgVisible "lblPwDesc2", 0
                    End If
                    DialogFunc = True
                Case "pwFromEH", "pwFromCST"
                    DialogFunc = True
                Case "Group3"
                    DlgVisible "lblRgnDesc0", 0
                    DlgVisible "lblRgnDesc1", 0
                    DlgVisible "lblRgnDesc2", 0
                    Select Case SuppValue
                        Case REGION_FORWARD
                            DlgVisible "lblRgnDesc1", 1
                        Case REGION_BACKWARD
                            DlgVisible "lblRgnDesc2", 1
                        Case Else
                            DlgVisible "lblRgnDesc0", 1
                    End Select
                    DialogFunc = True
                Case "about"
                    ShowAboutDialog
                    DialogFunc = True
                Case "btnCancel"
                    MsgBox "User cancelled the macro setup!" & vbCrLf & vbCrLf & _
                        "The enable flag is set to 0; the macro will not run!" & vbCrLf & vbCrLf & _
                        "But the CST sweep will continue!", _
                        vbCritical, "Parameter setup cancelled"
                    DialogFunc = False
                Case "help"
                    ShowHelpDialog
                    DialogFunc = True
                Case Else
                    DialogFunc = False
            End Select
        Case 4
    End Select
End Function

Private Function ShowParamsDialog(ByRef pos() As Double, _
    ByRef bAllModes As Boolean) As Boolean

    Begin Dialog UserDialog 1080, 578, "CST Eigenmode Slow-Wave Structure User Watch Macro  version " & g_sVersionString & " (" & g_sVersionLabel & ")" , .DialogFunc

        GroupBox 20, 7, 1040, 153, "[ Step 1 ] Select the modes for the interaction impedance calculation   ---   Number of eigenmodes requested from the solver: " & g_iNumModes & " modes"

        OptionGroup .Group1
            OptionButton 60, 28,  150, 14, " Compute all modes", .allMode
            OptionButton 60, 50,  200, 14, " Compute selected modes only", .someMode
            OptionButton 60, 118, 130, 14, " Custom modes", .someMode2

        CheckBox 50,  72, 80, 18, "Mode 1",  .selectMode1
        CheckBox 150, 72, 80, 18, "Mode 2",  .selectMode2
        CheckBox 250, 72, 80, 18, "Mode 3",  .selectMode3
        CheckBox 350, 72, 80, 18, "Mode 4",  .selectMode4
        CheckBox 450, 72, 80, 18, "Mode 5",  .selectMode5
        CheckBox 550, 72, 80, 18, "Mode 6",  .selectMode6
        CheckBox 650, 72, 80, 18, "Mode 7",  .selectMode7
        CheckBox 750, 72, 80, 18, "Mode 8",  .selectMode8
        CheckBox 850, 72, 80, 18, "Mode 9",  .selectMode9
        CheckBox 950, 72, 80, 18, "Mode 10", .selectMode10
        CheckBox 50,  94, 80, 18, "Mode 11", .selectMode11
        CheckBox 150, 94, 80, 18, "Mode 12", .selectMode12
        CheckBox 250, 94, 80, 18, "Mode 13", .selectMode13
        CheckBox 350, 94, 80, 18, "Mode 14", .selectMode14
        CheckBox 450, 94, 80, 18, "Mode 15", .selectMode15
        CheckBox 550, 94, 80, 18, "Mode 16", .selectMode16
        CheckBox 650, 94, 80, 18, "Mode 17", .selectMode17
        CheckBox 750, 94, 80, 18, "Mode 18", .selectMode18
        CheckBox 850, 94, 80, 18, "Mode 19", .selectMode19
        CheckBox 950, 94, 80, 18, "Mode 20", .selectMode20

        TextBox 200, 116, 830, 18, .txtMode
        Text    388, 138, 320, 14, "", .lblHint

        GroupBox 20, 167, 1040, 69, "[ Step 2 ] Set the interaction impedance reference point"

        Text  60, 188, 560, 14, "Default reference point coordinates   X = 0  Y = 0  Z = 0   units: " & g_sUnit

        Text     60, 212, 140, 14, "Reference X coordinate"
        TextBox 210, 210, 140, 18, .KcPosx
        Text    400, 212, 140, 14, "Reference Y coordinate"
        TextBox 550, 210, 140, 18, .KcPosy
        Text    740, 212, 140, 14, "Reference Z coordinate"
        TextBox 890, 210, 140, 18, .KcPosz

        GroupBox 20, 243, 1040, 65, "[ Step 3 ] Choose the power-flow calculation method"

        OptionGroup .Group2
            OptionButton 60, 264, 320, 14, "Compute the power flow from the E and H fields", .pwFromEH
            OptionButton 60, 286, 320, 14, "Use the native CST power-flow field", .pwFromCST

        Text 400, 264, 420, 14, "Indirect: interpolate, then cross-product, then integrate (approximate product)", .lblPwDesc1
        Text 400, 286, 420, 14, "Native: cross-product, then interpolate, then integrate (integrates the product)", .lblPwDesc2

        GroupBox 20, 315, 1040, 87, "[ Step 4 ] Choose the region whose interaction impedance is calculated"

        OptionGroup .Group3
            OptionButton 60, 336, 320, 14, "Compute the forward- and backward-wave regions", .regionBoth
            OptionButton 60, 358, 320, 14, "Compute the forward-wave region only", .regionFwd
            OptionButton 60, 380, 320, 14, "Compute the backward-wave region only", .regionBwd

        Text 400, 336, 640, 14, "No frequency direction check: every sweep point is used", .lblRgnDesc0
        Text 400, 358, 640, 14, "Forward-wave region: only points whose frequency rises with phase are used", .lblRgnDesc1
        Text 400, 380, 640, 14, "Backward-wave region: only points whose frequency falls with phase are used", .lblRgnDesc2

        GroupBox 20, 409, 1040, 109, "[ Notes ]"

        Text 60, 430, 990, 14, "1. The macro only checks whether Macro_Modex is defined in the parameter list to decide if the matching mode takes part in the interaction impedance calculation; the actual value does not affect the decision."
        Text 60, 452, 990, 14, "2. This dialog appears on the first run so that the user can choose the modes; afterwards the macro matches them automatically by detecting Macro_Modex in the parameter list."
        Text 60, 474, 990, 14, "3. The integer part of Macro_SweepWatch_Enable selects the power-flow method and its decimal digit selects the region: x.0 both regions, x.1 forward-wave region, x.2 backward-wave region."
        Text 60, 496, 990, 14, "4. Delete any one of Kc_RefPos_x/y/z, Macro_Modex or Macro_SweepWatch_Enable to make this dialog appear again."

        PushButton    20, 526, 90,  42, "About", .about
        CancelButton 120, 526, 90,  42, .btnCancel
        PushButton   870, 526, 90,  42, "Help", .help
        OKButton     970, 526, 90,  42

        Text 468, 540, 144, 14,"CST Macro Parameter Setup"

    End Dialog

    Dim dlg As UserDialog
    dlg.Group1 = 0
    dlg.Group2 = g_iPowerFlowSrc
    dlg.Group3 = g_iRegionMode
    dlg.KcPosx = CStr(pos(1))
    dlg.KcPosy = CStr(pos(2))
    dlg.KcPosz = CStr(pos(3))

    Dim nResult As Integer
    Dim bConfirmed As Boolean
    bConfirmed = False

    Do While Not bConfirmed
        Do
            nResult = Dialog(dlg)
            If nResult = 0 Then
                ShowParamsDialog = False
                Exit Function
            End If
        Loop While nResult = 1

        Dim selModes As String
        selModes = ""

        Select Case dlg.Group1
            Case 0
                g_bAllModes = True
            Case 1
                g_bAllModes = False
                If dlg.selectMode1  Then selModes = selModes & "1,"
                If dlg.selectMode2  Then selModes = selModes & "2,"
                If dlg.selectMode3  Then selModes = selModes & "3,"
                If dlg.selectMode4  Then selModes = selModes & "4,"
                If dlg.selectMode5  Then selModes = selModes & "5,"
                If dlg.selectMode6  Then selModes = selModes & "6,"
                If dlg.selectMode7  Then selModes = selModes & "7,"
                If dlg.selectMode8  Then selModes = selModes & "8,"
                If dlg.selectMode9  Then selModes = selModes & "9,"
                If dlg.selectMode10 Then selModes = selModes & "10,"
                If dlg.selectMode11 Then selModes = selModes & "11,"
                If dlg.selectMode12 Then selModes = selModes & "12,"
                If dlg.selectMode13 Then selModes = selModes & "13,"
                If dlg.selectMode14 Then selModes = selModes & "14,"
                If dlg.selectMode15 Then selModes = selModes & "15,"
                If dlg.selectMode16 Then selModes = selModes & "16,"
                If dlg.selectMode17 Then selModes = selModes & "17,"
                If dlg.selectMode18 Then selModes = selModes & "18,"
                If dlg.selectMode19 Then selModes = selModes & "19,"
                If dlg.selectMode20 Then selModes = selModes & "20,"
                If Len(selModes) > 0 Then selModes = Left(selModes, Len(selModes) - 1)
            Case 2
                g_bAllModes = False
                selModes = Trim(dlg.txtMode)
        End Select

        If dlg.Group2 = 1 Then
            g_iPowerFlowSrc = 1
        Else
            g_iPowerFlowSrc = 0
        End If

        Select Case dlg.Group3
            Case REGION_FORWARD
                g_iRegionMode = REGION_FORWARD
            Case REGION_BACKWARD
                g_iRegionMode = REGION_BACKWARD
            Case Else
                g_iRegionMode = REGION_BOTH
        End Select

        If g_sDirLabel = "X" Then
            pos(1) = 0#
        Else
            pos(1) = SafeCDbl(dlg.KcPosx, pos(1))
        End If
        If g_sDirLabel = "Y" Then
            pos(2) = 0#
        Else
            pos(2) = SafeCDbl(dlg.KcPosy, pos(2))
        End If
        If g_sDirLabel = "Z" Then
            pos(3) = 0#
        Else
            pos(3) = SafeCDbl(dlg.KcPosz, pos(3))
        End If

        If Not g_bAllModes And Len(selModes) = 0 Then
            If dlg.Group1 = 2 Then
                MsgBox "Enter at least one mode!", vbExclamation, "Note"
            Else
                MsgBox "Tick at least one mode!", vbExclamation, "Note"
            End If
        Else
            If g_bAllModes Then
                g_nSelectedModes = 0
            Else
                Dim parts() As String
                selModes = Replace(selModes, ", ", ",")
                selModes = Replace(selModes, "、", ",")
                selModes = Replace(selModes, " ", ",")
                parts = Split(selModes, ",")
                g_nSelectedModes = 0
                ReDim g_aSelectedModes(0 To UBound(parts))
                Dim j As Integer
                Dim k As Integer
                Dim nRaw As Integer
                Dim dModeTmp As Double
                Dim nModeIn As Long
                Dim sItem As String
                nRaw = 0
                For j = 0 To UBound(parts)
                    sItem = Trim(parts(j))

                    If Len(sItem) = 0 Then
                    ElseIf Not IsNumeric(sItem) Then
                        LogInfo "Note: the custom mode entry """ & sItem & """ is not a number; ignored"
                    Else
                        dModeTmp = SafeCDbl(sItem, 0)
                        If dModeTmp <> Int(dModeTmp) Then
                            LogInfo "Note: the custom mode entry """ & sItem & """ is not an integer; ignored"
                        ElseIf dModeTmp < 1# Or dModeTmp > CDbl(g_iNumModes) Then
                            LogInfo "Note: mode " & sItem & " is outside the solver mode range (1 - " & _
                                g_iNumModes & "); ignored"
                        Else
                            nRaw = nRaw + 1
                            nModeIn = CLng(dModeTmp)
                            k = 0
                            Do While k < g_nSelectedModes
                                If g_aSelectedModes(k) = nModeIn Then Exit Do
                                k = k + 1
                            Loop
                            If k = g_nSelectedModes Then
                                g_aSelectedModes(g_nSelectedModes) = CInt(nModeIn)
                                g_nSelectedModes = g_nSelectedModes + 1
                            End If
                        End If
                    End If
                Next j
                If g_nSelectedModes = 0 Then
                    g_bAllModes = True
                Else
                    ReDim Preserve g_aSelectedModes(0 To g_nSelectedModes - 1)
                    SortModeSelection g_aSelectedModes, g_nSelectedModes
                    If nRaw > g_nSelectedModes Then
                        LogInfo "Note: duplicate mode numbers in the input were removed automatically (" & _
                            nRaw & " -> " & g_nSelectedModes & ")"
                    End If
                    TruncateModeSelection g_aSelectedModes, g_nSelectedModes, _
                        g_bAllModes, g_iNumModes
                End If
            End If

            If Not g_bAllModes Then
                selModes = ""
                For j = 0 To g_nSelectedModes - 1
                    If j > 0 Then selModes = selModes & ", "
                    selModes = selModes & CStr(g_aSelectedModes(j))
                Next j
            End If

            Dim sConfirmMsg As String
            sConfirmMsg = "Please confirm the following settings: " & vbCrLf & vbCrLf
            If g_bAllModes Then
                sConfirmMsg = sConfirmMsg & "Modes: all (1 - " & g_iNumModes & ")" & vbCrLf & vbCrLf
            Else
                sConfirmMsg = sConfirmMsg & "Modes: " & selModes & vbCrLf & vbCrLf
            End If
            sConfirmMsg = sConfirmMsg & "Reference X: " & pos(1) & " " & g_sUnit & vbCrLf
            sConfirmMsg = sConfirmMsg & "Reference Y: " & pos(2) & " " & g_sUnit & vbCrLf
            sConfirmMsg = sConfirmMsg & "Reference Z: " & pos(3) & " " & g_sUnit & vbCrLf & vbCrLf
            sConfirmMsg = sConfirmMsg & "Power-flow calculation: using the "& PowerFlowSrcText(g_iPowerFlowSrc) & vbCrLf
            sConfirmMsg = sConfirmMsg & "Region: " & RegionText(g_iRegionMode) & vbCrLf & vbCrLf
            sConfirmMsg = sConfirmMsg & "Continue?"

            If MsgBox(sConfirmMsg, vbYesNo + vbQuestion, "Confirm settings") = vbYes Then
                bConfirmed = True
            End If
        End If
    Loop

    bAllModes = g_bAllModes
    ShowParamsDialog = True
End Function

Private Sub ShowAboutDialog()

    Begin Dialog UserDialog 1160, 370, "CST Eigenmode Slow-Wave Structure User Watch Macro --- About  version " & g_sVersionString & " (" & g_sVersionLabel & ")"

        Text 416, 20, 328, 14, "Normalized phase velocity / interaction impedance / Brillouin diagram calculation"
        Text  48, 42, 660, 14, "[ Author ]  " & g_sAuthorName
        Text  48, 64, 1000, 14, "[ Release date ]  " & g_sReleaseDate & "      Platform: " & g_sPlatform

        GroupBox 20, 86, 1120, 87, "[ Disclaimer ]"

        Text 48, 107, 1072, 14, "1. The author is responsible for the technical implementation of the macro."
        Text 48, 129, 1072, 14, "2. Released under the MIT License: free to use, modify, redistribute and use commercially, provided the original copyright and license notice are kept (see the accompanying LICENSE)."
        Text 48, 151, 1072, 14, "3. The code is provided ""as is"" without any guarantee of the absolute accuracy of the results; users must verify critical simulation results themselves, and the author accepts no liability arising from its use."

        GroupBox 20, 181, 1120, 131, "[ Prerequisites ]"

        Text 48, 202, 1072, 14, "1. Use the tetrahedral or hexahedral mesh algorithm; the hexahedral JDM algorithm is recommended because it is accurate."
        Text 48, 224, 1072, 14, "2. Periodic boundary conditions are set on the slow-wave structure."
        Text 48, 246, 1072, 14, "3. The parameter sweep list must contain the phase parameter."
        Text 48, 268, 1072, 14, "4. Slow Wave Userdefined Watch has been added to the CST macro commands."
        Text 48, 290, 1072, 14, "5. The CST version is greater than or equal to " & g_sCstVerMin & " and not greater than " & g_sCstVerMax & ", and the core computation library LinXi.dll is deployed correctly."

        OKButton 530, 320, 100, 42

    End Dialog

    Dim dlg As UserDialog
    Dialog dlg

End Sub

Private Sub ShowHelpDialog()

    Begin Dialog UserDialog 1160, 645, "CST Eigenmode Slow-Wave Structure User Watch Macro --- Help  version " & g_sVersionString & " (" & g_sVersionLabel & ")"

        Text 428, 20, 264, 14, "CST Eigenmode Sweep Macro Help"

        GroupBox 20, 42, 520, 295, "[ Parameters ]"

        Text  36,  63, 484, 14, "1. Macro_SweepWatch_Enable = macro enable flag + power-flow calculation method + region"
        Text  56,  91, 464, 14, "， integer part: 1 = indirect power-flow method; 2 = native power-flow method; 0 = disabled."
        Text  56, 119, 464, 14, "， decimal digit: 0 = both regions; 1 = forward-wave region only; 2 = backward-wave region only."
        Text  56, 147, 464, 14, "， the forward-wave region requires the frequency to rise with phase, the backward-wave region requires it to fall."
        Text  56, 175, 464, 14, "， any other value (for example 1.3 or 1.11) makes the parameter setup dialog appear again on the next run."
        Text  36, 203, 484, 14, "2. Kc_RefPos_x / y / z = 3D coordinates of the interaction impedance reference point"
        Text  56, 231, 464, 14, "， the periodic direction is locked to the centre of the computation domain and cannot be changed."
        Text  56, 259, 464, 14, "， the non-periodic directions can be set freely by the user."
        Text  36, 287, 484, 14, "3. Macro_Modex (x = 1, 2, 3 ...) = interaction impedance calculation flag"
        Text  56, 315, 464, 14, "， the macro only checks whether this parameter exists in the parameter list, not its value."

        GroupBox 560, 42, 580, 295, "[ Parameter setup and logging ]"

        Text 576,  63, 548, 14, "-- Parameter setup"
        Text 576,  91, 548, 14, "， tick the modes or type the mode numbers as the dialog prompts."
        Text 576, 119, 548, 14, "， mode numbers beyond the solver's eigenmode count are truncated automatically."
        Text 576, 147, 548, 14, "， two power-flow calculation methods can be selected."
        Text 576, 175, 548, 14, "， the region can be the forward-wave one, the backward-wave one, or both."
        Text 576, 203, 548, 14, "， with both regions selected, the frequency direction is not checked at all."
        Text 576, 231, 548, 14, "-- Logging"
        Text 576, 259, 548, 14, "， log file: project directory \Temp\macro_log.txt"
        Text 576, 287, 548, 14, "， records the whole sweep (progress and error messages) to help with troubleshooting."

        GroupBox 20, 347, 1120, 239, "[ Important notes ]"

        Text  36, 368, 528, 14, "， the JDM algorithm is recommended for the solver because it is accurate."
        Text 592, 368, 528, 14, "， with the JDM algorithm, eigenmode names must not contain a decimal point."
        Text  36, 396, 528, 14, "， do not operate the current window manually during the sweep."
        Text 592, 396, 528, 14, "， clear the existing results before repeating a sweep, otherwise the calculated results are lost."
        Text  36, 424, 528, 14, "， one window, one task; use several CST windows for several tasks."
        Text 592, 424, 528, 14, "， to pause, just cancel the sweep task."
        Text  36, 452, 1088, 14, "， deleting the slow-wave structure user-watch step from the history tree does not remove the macro itself."
        Text  36, 480, 1088, 14, "， do not open the Fields on Plane or Cutting Plane views of the electric or magnetic field in an eigenmode project."
        Text  36, 508, 1088, 14, "， the parameters the macro relies on - Macro_SweepWatch_Enable, Macro_Modex and the periodic-direction Kc_RefPos - must not be swept."
        Text  36, 536, 1088, 14, "， deleting Macro_SweepWatch_Enable makes the parameter dialog appear again on the next run so that the power-flow method and the region can be chosen again."
        Text  36, 564, 1088, 14, "， keep the project in a short directory so that the path stays inside the Windows 260-character limit and the data files can be read and written."

        OKButton 530, 595, 100, 42

    End Dialog

    Dim dlg As UserDialog
    Dialog dlg

End Sub

Private Function SafeCDbl(ByVal sVal As String, ByVal dDefault As Double) As Double
    Err.Clear
    On Error Resume Next
    SafeCDbl = CDbl(Trim(sVal))
    If Err <> 0 Then SafeCDbl = dDefault
    Err.Clear
    On Error GoTo 0
End Function

Private Function PowerFlowSrcText(ByVal iSrc As Integer) As String
    If iSrc = 1 Then
        PowerFlowSrcText = "native power-flow method"
    Else
        PowerFlowSrcText = "indirect power-flow method"
    End If
End Function

Private Function RegionText(ByVal iRegion As Integer) As String
    Select Case iRegion
        Case REGION_FORWARD
            RegionText = "forward-wave region"
        Case REGION_BACKWARD
            RegionText = "backward-wave region"
        Case Else
            RegionText = "forward- and backward-wave regions"
    End Select
End Function

Private Function EnableFlagText(ByVal iSrc As Integer, ByVal iRegion As Integer) As String
    EnableFlagText = "Macro activation state, using the " & PowerFlowSrcText(iSrc) & _
        ", computing the " & RegionText(iRegion)
End Function

Private Function DecodeEnableFlag(ByVal dValue As Double, _
    ByRef iSrc As Integer, ByRef iRegion As Integer) As Boolean

    Dim dScaled As Double
    Dim nCode As Long
    Dim nMethod As Long
    Dim nRegionCode As Long

    DecodeEnableFlag = False
    iSrc = 0
    iRegion = REGION_BOTH

    If dValue <= 0# Then Exit Function
    dScaled = dValue * CDbl(ENABLE_FLAG_SCALE)
    If dScaled > 1000000# Then Exit Function
    nCode = CLng(Int(dScaled + 0.5))
    If Abs(dScaled - CDbl(nCode)) > ENABLE_FLAG_EPS Then Exit Function

    nMethod = nCode \ ENABLE_FLAG_SCALE
    nRegionCode = nCode Mod ENABLE_FLAG_SCALE

    If nMethod <> PW_SRC_FROM_EH And nMethod <> PW_SRC_FROM_CST Then Exit Function
    If nRegionCode < REGION_BOTH Or nRegionCode > REGION_BACKWARD Then Exit Function

    iSrc = CInt(nMethod - 1)
    iRegion = CInt(nRegionCode)
    DecodeEnableFlag = True
End Function

Private Function HasIllegalChar(ByVal s As String) As Boolean
    HasIllegalChar = (InStr(s, "\") > 0 Or InStr(s, "/") > 0 Or InStr(s, ":") > 0 Or _
        InStr(s, "*") > 0 Or InStr(s, "?") > 0 Or InStr(s, """") > 0 Or _
        InStr(s, "<") > 0 Or InStr(s, ">") > 0 Or InStr(s, "|") > 0)
End Function

Private Function SafeNamePart(ByVal s As String) As String
    Dim i As Long, ch As String, sOut As String

    s = Replace(s, " & ", "+")
    s = Replace(s, " = ", "-")

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "\", "/", ":", "*", "?", """", "<", ">", "|"
                ch = "_"
        End Select
        sOut = sOut & ch
    Next i

    Do While InStr(sOut, "__") > 0
        sOut = Replace(sOut, "__", "_")
    Loop
    Do While Len(sOut) > 0
        ch = Right$(sOut, 1)
        If ch = "." Or ch = " " Then
            sOut = Left$(sOut, Len(sOut) - 1)
        Else
            Exit Do
        End If
    Loop

    SafeNamePart = sOut
End Function

Private Function MakeFileTag(ByVal sKey As String, ByVal nIndex As Long, _
    ByVal nBudget As Long) As String
    Dim sRead As String

    If Len(sKey) = 0 Then
        MakeFileTag = ""
        Exit Function
    End If

    If Len(sKey) <= nBudget And Not HasIllegalChar(sKey) Then
        MakeFileTag = sKey
        Exit Function
    End If

    If nBudget < 12 Then nBudget = 12
    sRead = SafeNamePart(sKey)
    If Len(sRead) > nBudget - 6 Then sRead = Left$(sRead, nBudget - 6)

    MakeFileTag = sRead & "_g" & CStr(nIndex)
End Function

Private Function LookupGroupIndex(ByVal sKey As String, ByRef nCount As Long) As Long
    Dim f As Integer
    Dim sLine As String, sK As String, sIdx As String
    Dim p1 As Long, p2 As Long, nIdx As Long

    nCount = 0
    LookupGroupIndex = 0
    If Len(sKey) = 0 Then Exit Function
    If Dir(g_sTemp & "_groups.txt") = "" Then Exit Function

    Err.Clear
    On Error Resume Next
    f = FreeFile
    Open g_sTemp & "_groups.txt" For Input As #f
    If Err <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Do While Not EOF(f)
        Line Input #f, sLine
        If Trim(sLine) <> "" Then
            p1 = InStr(sLine, "|")
            If p1 > 0 Then
                sK = Mid(sLine, p1 + 1)
                p2 = InStr(sK, "|")
                If p2 > 0 Then
                    sIdx = Mid(sK, p2 + 1)
                    sK = Left(sK, p2 - 1)
                Else
                    sIdx = ""
                End If
            Else
                sK = sLine
                sIdx = ""
            End If

            nIdx = CLng(Val(sIdx))
            If nIdx > nCount Then nCount = nIdx
            If nIdx > 0 And sK = sKey Then
                LookupGroupIndex = nIdx
                Exit Do
            End If
        End If
    Loop
    Close #f
    Err.Clear
    On Error GoTo 0
End Function

Private Sub BuildModeFlagTable()
    Dim i As Integer
    ReDim g_abModeFlag(1 To g_iNumModes)
    If g_bAllModes Then
        For i = 1 To g_iNumModes
            g_abModeFlag(i) = True
        Next i
    Else
        For i = 1 To g_iNumModes
            g_abModeFlag(i) = IsModeInSelection(i, g_aSelectedModes, g_nSelectedModes)
        Next i
    End If
End Sub

Private Sub DeleteTreeItemRecursive(ByVal sPath As String)
    If Not ResultTree.DoesTreeItemExist(sPath) Then Exit Sub
    Dim sChild As String, sNext As String
    sChild = ResultTree.GetFirstChildName(sPath)
    Do While sChild <> ""
        sNext = ResultTree.GetNextItemName(sChild)
        DeleteTreeItemRecursive sChild
        sChild = sNext
    Loop
    If ResultTree.DoesTreeItemExist(sPath) Then
        With ResultTree
            .Name sPath
            .Delete
        End With
    End If
End Sub

Private Sub DeleteFilesByPattern(ByVal sPattern As String)
    Dim sDir As String, sFile As String
    On Error Resume Next

    sDir = Left(sPattern, InStrRev(sPattern, "\"))
    sFile = Dir(sPattern)
    Do While sFile <> ""
        Kill sDir & sFile
        sFile = Dir()
    Loop

    On Error GoTo 0
End Sub

Private Sub SortModeSelection(ByRef aModes() As Integer, ByVal nCount As Integer)
    Dim i As Integer, j As Integer, nTmp As Integer
    If nCount < 2 Then Exit Sub
    For i = 0 To nCount - 2
        For j = 0 To nCount - 2 - i
            If aModes(j) > aModes(j + 1) Then
                nTmp = aModes(j)
                aModes(j) = aModes(j + 1)
                aModes(j + 1) = nTmp
            End If
        Next j
    Next i
End Sub

Private Sub TruncateModeSelection(ByRef aModes() As Integer, _
    ByRef nCount As Integer, ByRef bAll As Boolean, _
    ByVal nMaxMode As Integer)

    If bAll Then Exit Sub

    Dim i As Integer
    Dim bTruncated As Boolean
    bTruncated = False

    Dim nValid As Integer
    nValid = 0
    For i = 0 To nCount - 1
        If aModes(i) <= nMaxMode Then
            nValid = nValid + 1
        Else
            bTruncated = True
        End If
    Next i

    If Not bTruncated Then Exit Sub

    If nValid = 0 Then
        bAll = True
        nCount = 0
        LogInfo "Warning: all selected modes are out of range (max Mode = " & nMaxMode & "); switching to all modes"
    Else
        Dim aValid() As Integer
        ReDim aValid(0 To nValid - 1)
        Dim j As Integer
        j = 0
        For i = 0 To nCount - 1
            If aModes(i) <= nMaxMode Then
                aValid(j) = aModes(i)
                j = j + 1
            End If
        Next i
        nCount = nValid
        ReDim aModes(0 To nValid - 1)
        For i = 0 To nValid - 1
            aModes(i) = aValid(i)
        Next i
        LogInfo "Warning: the selection contained out-of-range mode numbers (max Mode = " & nMaxMode & "); truncated automatically"
    End If
End Sub

Private Function IsModeInSelection(ByVal iMode As Integer, _
    ByRef aModes() As Integer, ByVal nCount As Integer) As Boolean
    Dim i As Integer
    For i = 0 To nCount - 1
        If aModes(i) = iMode Then
            IsModeInSelection = True
            Exit Function
        End If
    Next i
    IsModeInSelection = False
End Function

Private Sub SwitchToItem(ByVal sPath As String)
    On Error Resume Next
    SelectTreeItem sPath
    On Error GoTo 0
    DoEvents
    Wait WAIT_SEC
End Sub

Private Sub Save1D(ByRef oRes As Object, ByVal sTitle As String, ByVal sPath As String, ByVal sTree As String, _
    Optional ByVal sXLabel As String = "Frequency (GHz)", _
    Optional ByVal sYLabel As String = "")
    With oRes
        .Title sTitle
        .XLabel sXLabel
        If Len(sYLabel) > 0 Then .YLabel sYLabel
        .Type "Linear"
        .Save sPath
        .AddToTree sTree
    End With
End Sub

Private Sub FailGrid(ByVal sMsg As String)
    LogError sMsg
    g_n1 = 0
    g_n2 = 0
    g_nL = 0
    g_nCross = 0
    g_bGridFailed = True
End Sub

Private Sub PrecomputeMeshGrid()
    If g_nCross > 0 Then Exit Sub
    If g_bGridFailed Then Exit Sub

    If g_bIsTetra Then
        PrecomputeRegularGrid
        Exit Sub
    End If

    Dim iw1Low As Long, iw1High As Long, iw2Low As Long, iw2High As Long
    Dim iwLow As Long, iwHigh As Long
    Dim i1 As Long, i2 As Long, iw As Long

    Select Case g_iDir
        Case 1
            iw1Low = Mesh.GetClosestYIndex(CStr(g_cMin(2))): iw1High = Mesh.GetClosestYIndex(CStr(g_cMax(2)))
            iw2Low = Mesh.GetClosestZIndex(CStr(g_cMin(3))): iw2High = Mesh.GetClosestZIndex(CStr(g_cMax(3)))
        Case 2
            iw1Low = Mesh.GetClosestXIndex(CStr(g_cMin(1))): iw1High = Mesh.GetClosestXIndex(CStr(g_cMax(1)))
            iw2Low = Mesh.GetClosestZIndex(CStr(g_cMin(3))): iw2High = Mesh.GetClosestZIndex(CStr(g_cMax(3)))
        Case 3
            iw1Low = Mesh.GetClosestXIndex(CStr(g_cMin(1))): iw1High = Mesh.GetClosestXIndex(CStr(g_cMax(1)))
            iw2Low = Mesh.GetClosestYIndex(CStr(g_cMin(2))): iw2High = Mesh.GetClosestYIndex(CStr(g_cMax(2)))
    End Select

    g_n1 = iw1High - iw1Low
    g_n2 = iw2High - iw2Low
    If g_n1 <= 0 Or g_n2 <= 0 Then
        FailGrid "Power-flow cross-section: the grid count is 0 (n1=" & g_n1 & ", n2=" & g_n2 & ")"
        Exit Sub
    End If

    Select Case g_iDir
        Case 1: iwLow = Mesh.GetClosestXIndex(CStr(g_cMin(1))): iwHigh = Mesh.GetClosestXIndex(CStr(g_cMax(1)))
        Case 2: iwLow = Mesh.GetClosestYIndex(CStr(g_cMin(2))): iwHigh = Mesh.GetClosestYIndex(CStr(g_cMax(2)))
        Case 3: iwLow = Mesh.GetClosestZIndex(CStr(g_cMin(3))): iwHigh = Mesh.GetClosestZIndex(CStr(g_cMax(3)))
    End Select

    g_nL = iwHigh - iwLow
    If g_nL <= 0 Then
        FailGrid "Longitudinal integration: the grid count is 0"
        Exit Sub
    End If

    g_nCross = g_n1 * g_n2

    Dim coord1() As Double, coord2() As Double, coordL() As Double
    ReDim coord1(0 To g_n1): ReDim coord2(0 To g_n2): ReDim coordL(0 To g_nL)

    For i1 = 0 To g_n1
        Select Case g_iDir
            Case 1: coord1(i1) = Mesh.GetY(iw1Low + i1)
            Case 2, 3: coord1(i1) = Mesh.GetX(iw1Low + i1)
        End Select
    Next i1
    For i2 = 0 To g_n2
        Select Case g_iDir
            Case 3: coord2(i2) = Mesh.GetY(iw2Low + i2)
            Case 1, 2: coord2(i2) = Mesh.GetZ(iw2Low + i2)
        End Select
    Next i2
    For iw = 0 To g_nL
        Select Case g_iDir
            Case 1: coordL(iw) = Mesh.GetX(iwLow + iw)
            Case 2: coordL(iw) = Mesh.GetY(iwLow + iw)
            Case 3: coordL(iw) = Mesh.GetZ(iwLow + iw)
        End Select
    Next iw

    Dim arrMid1() As Double, arrMid2() As Double
    ReDim arrMid1(0 To g_n1 - 1): ReDim g_arrDw1(0 To g_n1 - 1)
    ReDim arrMid2(0 To g_n2 - 1): ReDim g_arrDw2(0 To g_n2 - 1)
    ReDim g_arrMidL(0 To g_nL - 1): ReDim g_arrDwL(0 To g_nL - 1)

    Dim nRet As Long
    nRet = CBuildHexaGrid(CORE_AUTHOR, g_iDir, g_n1, g_n2, g_nL, _
        coord1(0), coord2(0), coordL(0), _
        g_dUnitToSI, g_xRef, g_yRef, g_zRef, _
        g_sPointsFile, _
        arrMid1(0), g_arrDw1(0), arrMid2(0), g_arrDw2(0), g_arrMidL(0), g_arrDwL(0))
    If nRet <> 0 Then
        FailGrid "Hexahedral sampling DLL failed (return code " & nRet & ")"
        Exit Sub
    End If

    Select Case g_iDir
        Case 1
            LogInfo "  Sampling grid: periodic direction X = " & g_nL & _
                ", transverse direction Y = " & g_n1 & ", transverse direction Z = " & g_n2
        Case 2
            LogInfo "  Sampling grid: periodic direction Y = " & g_nL & _
                ", transverse direction X = " & g_n1 & ", transverse direction Z = " & g_n2
        Case 3
            LogInfo "  Sampling grid: periodic direction Z = " & g_nL & _
                ", transverse direction X = " & g_n1 & ", transverse direction Y = " & g_n2
    End Select

End Sub

Private Sub DetermineTetraStep()
    Dim dLen1 As Double, dLen2 As Double, dLenL As Double
    g_n1 = 0: g_n2 = 0: g_nL = 0
    Select Case g_iDir
        Case 1: dLen1 = g_cMax(2) - g_cMin(2): dLen2 = g_cMax(3) - g_cMin(3): dLenL = g_cMax(1) - g_cMin(1)
        Case 2: dLen1 = g_cMax(1) - g_cMin(1): dLen2 = g_cMax(3) - g_cMin(3): dLenL = g_cMax(2) - g_cMin(2)
        Case 3: dLen1 = g_cMax(1) - g_cMin(1): dLen2 = g_cMax(2) - g_cMin(2): dLenL = g_cMax(3) - g_cMin(3)
    End Select

    Dim dMinEdge As Double, dMaxEdge As Double
    dMinEdge = Mesh.GetMinimumEdgeLength
    dMaxEdge = Mesh.GetMaximumEdgeLength

    Dim n1 As Long, n2 As Long, nL As Long
    Dim h As Double
    Dim nRet As Long
    nRet = CComputeTetraStep(CORE_AUTHOR, dLen1, dLen2, dLenL, dMinEdge, dMaxEdge, _
        MAX_CROSS_POINTS, MAX_PTS_PER_DIR, n1, n2, nL, h)
    If nRet <> 0 Then
        FailGrid "Tetrahedral step-size DLL failed (return code " & nRet & ")"
        Exit Sub
    End If
    If n1 < 1 Or n2 < 1 Or nL < 1 Then
        FailGrid "Tetrahedral step-size DLL returned invalid dimensions (n1=" & n1 & ", n2=" & n2 & ", nL=" & nL & ")"
        Exit Sub
    End If

    g_n1 = n1: g_n2 = n2: g_nL = nL
    LogInfo "  Tetrahedral automatic step: h = " & Format(h, "0.000E+00") & " " & g_sUnit
End Sub

Private Sub BuildTetraGrid(ByVal n1 As Long, ByVal n2 As Long, ByVal nL As Long)
    g_n1 = n1: g_n2 = n2: g_nL = nL
    If g_n1 < 1 Or g_n2 < 1 Or g_nL < 1 Then
        FailGrid "Invalid tetrahedral sampling grid dimensions (n1=" & g_n1 & ", n2=" & g_n2 & _
            ", nL=" & g_nL & "); skipping this calculation"
        Exit Sub
    End If
    g_nCross = g_n1 * g_n2

    Dim cMin1 As Double, cMax1 As Double, cMin2 As Double, cMax2 As Double
    Dim cMinL As Double, cMaxL As Double
    Select Case g_iDir
        Case 1: cMin1 = g_cMin(2): cMax1 = g_cMax(2): cMin2 = g_cMin(3): cMax2 = g_cMax(3): cMinL = g_cMin(1): cMaxL = g_cMax(1)
        Case 2: cMin1 = g_cMin(1): cMax1 = g_cMax(1): cMin2 = g_cMin(3): cMax2 = g_cMax(3): cMinL = g_cMin(2): cMaxL = g_cMax(2)
        Case 3: cMin1 = g_cMin(1): cMax1 = g_cMax(1): cMin2 = g_cMin(2): cMax2 = g_cMax(2): cMinL = g_cMin(3): cMaxL = g_cMax(3)
    End Select

    Dim arrMid1() As Double, arrMid2() As Double
    ReDim arrMid1(0 To g_n1 - 1): ReDim g_arrDw1(0 To g_n1 - 1)
    ReDim arrMid2(0 To g_n2 - 1): ReDim g_arrDw2(0 To g_n2 - 1)
    ReDim g_arrMidL(0 To g_nL - 1): ReDim g_arrDwL(0 To g_nL - 1)

    Dim nRet As Long
    nRet = CBuildUniformGrid( _
        CORE_AUTHOR, _
        g_n1, g_n2, g_nL, _
        cMin1, cMax1, cMin2, cMax2, cMinL, cMaxL, _
        g_dUnitToSI, _
        arrMid1(0), g_arrDw1(0), arrMid2(0), g_arrDw2(0), g_arrMidL(0), g_arrDwL(0), _
        g_iDir, g_xRef, g_yRef, g_zRef, g_sPointsFile)
    If nRet <> 0 Then
        FailGrid "Tetrahedral grid-geometry DLL failed (return code " & nRet & ")"
        Exit Sub
    End If
End Sub

Private Sub PrecomputeRegularGrid()
    DetermineTetraStep
    If g_n1 < 1 Or g_n2 < 1 Or g_nL < 1 Then Exit Sub
    BuildTetraGrid g_n1, g_n2, g_nL
    If g_nCross < 1 Then Exit Sub
    Select Case g_iDir
        Case 1
            LogInfo "  Sampling grid: periodic direction X = " & g_nL & _
                ", transverse direction Y = " & g_n1 & ", transverse direction Z = " & g_n2
        Case 2
            LogInfo "  Sampling grid: periodic direction Y = " & g_nL & _
                ", transverse direction X = " & g_n1 & ", transverse direction Z = " & g_n2
        Case 3
            LogInfo "  Sampling grid: periodic direction Z = " & g_nL & _
                ", transverse direction X = " & g_n1 & ", transverse direction Y = " & g_n2
    End Select

End Sub

Private Sub ExportFieldAtPoints(ByVal iMode As Integer, ByVal sField As String, ByVal sOutFile As String, _
    Optional ByVal bSkipSwitch As Boolean = False)
    On Error Resume Next
    Kill sOutFile
    On Error GoTo 0

    If Not bSkipSwitch Then
        SwitchToItem "2D/3D Results\Modes\Mode " & CStr(iMode) & "\" & sField
    End If

    With ASCIIExport
        .Reset
        .FileName sOutFile
        .SetPointFile g_sPointsFile
        .ExportCoordinatesInMeter False
        .Execute
    End With
End Sub

Private Function ParseFieldFile(ByVal sPath As String, _
    ByRef aExRe() As Double, ByRef aExIm() As Double, _
    ByRef aEyRe() As Double, ByRef aEyIm() As Double, _
    ByRef aEzRe() As Double, ByRef aEzIm() As Double) As Boolean
    Dim nExpected As Long
    nExpected = g_nCross + g_nL
    If nExpected <= 0 Then ParseFieldFile = False: Exit Function

    ReDim aExRe(0 To nExpected - 1): ReDim aExIm(0 To nExpected - 1)
    ReDim aEyRe(0 To nExpected - 1): ReDim aEyIm(0 To nExpected - 1)
    ReDim aEzRe(0 To nExpected - 1): ReDim aEzIm(0 To nExpected - 1)

    Dim dX0 As Double, dY0 As Double, dZ0 As Double
    Dim nRet As Long
    nRet = CParseFieldFile(CORE_AUTHOR, sPath, nExpected, dX0, dY0, dZ0, _
        aExRe(0), aExIm(0), aEyRe(0), aEyIm(0), aEzRe(0), aEzIm(0))

    If nRet = nExpected Then
        If Not g_bFirstFieldLog Then
            g_bFirstFieldLog = True
        End If
        ParseFieldFile = True
    Else
        LogError "Failed to parse the field file: " & sPath & " (return code " & nRet & ")"
        ParseFieldFile = False
    End If
End Function

Private Function ParsePowerFlowFile(ByVal sPath As String, _
    ByRef aPxRe() As Double, ByRef aPxIm() As Double, _
    ByRef aPyRe() As Double, ByRef aPyIm() As Double, _
    ByRef aPzRe() As Double, ByRef aPzIm() As Double) As Boolean
    Dim nExpected As Long
    nExpected = g_nCross + g_nL
    If nExpected <= 0 Then ParsePowerFlowFile = False: Exit Function

    ReDim aPxRe(0 To nExpected - 1): ReDim aPxIm(0 To nExpected - 1)
    ReDim aPyRe(0 To nExpected - 1): ReDim aPyIm(0 To nExpected - 1)
    ReDim aPzRe(0 To nExpected - 1): ReDim aPzIm(0 To nExpected - 1)

    Dim nRet As Long
    nRet = CParsePowerFlowFile(CORE_AUTHOR, sPath, nExpected, _
        aPxRe(0), aPxIm(0), aPyRe(0), aPyIm(0), aPzRe(0), aPzIm(0))

    If nRet = nExpected Then
        ParsePowerFlowFile = True
    Else
        LogError "Failed to parse the power-flow field file: " & sPath & " (return code " & nRet & ")"
        ParsePowerFlowFile = False
    End If
End Function

Private Function ExportEAndParse(ByVal iMode As Integer) As Boolean
    Dim sFile As String
    If g_nCross + g_nL <= 0 Then PrecomputeMeshGrid
    If g_nCross + g_nL <= 0 Then ExportEAndParse = False: Exit Function

    sFile = g_sTemp & "_E_mode" & iMode & ".txt"

    ExportFieldAtPoints iMode, "e", sFile, False
    If Not ParseFieldFile(sFile, g_aExRe, g_aExIm, g_aEyRe, g_aEyIm, g_aEzRe, g_aEzIm) Then
        ExportEAndParse = False
        Exit Function
    End If
    ExportEAndParse = True
End Function

Private Function ExportPAndParse(ByVal iMode As Integer) As Boolean
    Dim sFile As String
    sFile = g_sTemp & "_P_mode" & iMode & ".txt"
    ExportFieldAtPoints iMode, "Power Flow", sFile
    If Not ParsePowerFlowFile(sFile, g_aPxRe, g_aPxIm, g_aPyRe, g_aPyIm, g_aPzRe, g_aPzIm) Then
        ExportPAndParse = False
        Exit Function
    End If
    ExportPAndParse = True
End Function

Private Function ExportHAndParse(ByVal iMode As Integer) As Boolean
    Dim sFile As String
    If g_nCross + g_nL <= 0 Then PrecomputeMeshGrid
    If g_nCross + g_nL <= 0 Then ExportHAndParse = False: Exit Function

    sFile = g_sTemp & "_H_mode" & iMode & ".txt"
    ExportFieldAtPoints iMode, "h", sFile
    If Not ParseFieldFile(sFile, g_aHxRe, g_aHxIm, g_aHyRe, g_aHyIm, g_aHzRe, g_aHzIm) Then
        ExportHAndParse = False
        Exit Function
    End If
    ExportHAndParse = True
End Function

Private Function R3DGetFieldFrequency(ByVal sMode As String) As Double
    Dim res As Object
    Dim dFreq As Double
    Dim nErr As Long

    R3DGetFieldFrequency = 0#

    Err.Clear
    On Error Resume Next
    If g_bIsTetra Then
        Set res = Result3D("^e" & sMode & ".m3t")
    Else
        Set res = Result3D("^mode_e_" & sMode & ".m3d")
    End If
    nErr = Err.Number
    Err.Clear
    On Error GoTo 0

    If nErr <> 0 Or res Is Nothing Then
        LogWarning "  Mode " & sMode & ": frequency read failed, switching to the fallback method"
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    dFreq = res.GetFrequency()
    nErr = Err.Number
    Err.Clear
    On Error GoTo 0
    Set res = Nothing

    If nErr <> 0 Or dFreq <= 0# Then
        LogWarning "  Mode " & sMode & ": frequency read failed, switching to the fallback method"
        Exit Function
    End If

    R3DGetFieldFrequency = dFreq
End Function

Private Function LoadFieldDataFromEH(ByVal iMode As Integer) As Boolean
    If Not ExportEAndParse(iMode) Then
        LoadFieldDataFromEH = False
        Exit Function
    End If

    If Not ExportHAndParse(iMode) Then
        LoadFieldDataFromEH = False
        Exit Function
    End If

    LoadFieldDataFromEH = True
End Function

Private Function PreparePowerFieldFromCST(ByVal iMode As Integer, ByVal sMode As String) As Boolean
    PreparePowerFieldFromCST = False

    If Not ExportEAndParse(iMode) Then Exit Function

    Dim resEField As Object, resHField As Object
    Dim resHConj As Object, resPower As Object
    Dim nErr As Long

    Set resEField = Nothing: Set resHField = Nothing
    Set resHConj = Nothing:  Set resPower = Nothing

    Err.Clear
    On Error Resume Next
    If g_bIsTetra Then
        Set resEField = Result3D("^e" & sMode & ".m3t")
        Set resHField = Result3D("^h" & sMode & ".m3t")
    Else
        Set resEField = Result3D("^mode_e_" & sMode & ".m3d")
        Set resHField = Result3D("^mode_h_" & sMode & ".m3d")
    End If
    nErr = Err.Number
    Err.Clear
    On Error GoTo 0

    If nErr <> 0 Then
        LogError "  Mode " & sMode & ": failed to load the EM field data file (Err=" & nErr & ")"
    ElseIf resEField Is Nothing Or resHField Is Nothing Then
        LogError "  Mode " & sMode & ": failed to generate the power-flow vector field; the EM field data files were not found"
    ElseIf resEField.GetLength <= 0 Or resEField.GetLength <> resHField.GetLength Then
        LogError "  Mode " & sMode & ": failed to generate the power-flow vector field; the EM field data files may have failed to load or their point order may differ"
    Else
        Err.Clear
        On Error Resume Next
        Set resHConj = resHField.Copy
        resHConj.Conjugate
        Set resPower = resEField.Copy
        resPower.VectorProd resHConj
        resPower.setTitle "Poynting Vector"
        resPower.SetType "dynamic powerflow"
        If g_bIsTetra Then
            resPower.Save "^mode_p_" & sMode & ".m3t"
        Else
            resPower.Save "^mode_p_" & sMode & ".m3d"
        End If
        resPower.AddToTree "2D/3D Results\Modes\Mode " & sMode & "\Power Flow", ""
        nErr = Err.Number
        Err.Clear
        On Error GoTo 0

        If nErr <> 0 Then
            LogError "  Mode " & sMode & ": failed to generate / save the power-flow vector field (Err=" & nErr & ")"
        Else
            LogInfo "  Mode " & sMode & ": power-flow vector field generated successfully"

            If ExportPAndParse(iMode) Then
                PreparePowerFieldFromCST = True
            Else
                LogError "  Mode " & sMode & ": power-flow field export / parsing failed"
            End If
        End If
    End If

    Set resPower = Nothing
    Set resHConj = Nothing
    Set resEField = Nothing
    Set resHField = Nothing
End Function

Private Sub IntegrateLongitudinal(ByVal dBeta As Double, _
    ByRef dVre As Double, ByRef dVim As Double, ByRef dEabs As Double)
    If g_nL <= 0 Then
        dVre = 0#: dVim = 0#: dEabs = 0#
        Exit Sub
    End If
    Dim nRet As Long
    nRet = CIntegrateLongitudinal( _
        CORE_AUTHOR, _
        g_iDir, g_nL, g_nCross, _
        dBeta, g_dUnitToSI, g_dPitchSI, _
        g_arrMidL(0), g_arrDwL(0), _
        g_aExRe(0), g_aExIm(0), _
        g_aEyRe(0), g_aEyIm(0), _
        g_aEzRe(0), g_aEzIm(0), _
        dVre, dVim, dEabs)
    If nRet <> 0 Then
        LogError "Longitudinal integration DLL call failed (return code " & nRet & ")"
        dVre = 0#: dVim = 0#: dEabs = 0#
    End If
End Sub

Private Sub ComputePowerFlowFromCST(ByRef dPower As Double)
    If g_n1 <= 0 Or g_n2 <= 0 Then
        dPower = 0#
        Exit Sub
    End If
    dPower = CIntegratePowerFlow( _
        CORE_AUTHOR, _
        g_iDir, g_n1, g_n2, _
        g_arrDw1(0), g_arrDw2(0), _
        g_aPxRe(0), g_aPxIm(0), _
        g_aPyRe(0), g_aPyIm(0), _
        g_aPzRe(0), g_aPzIm(0))
End Sub

Private Sub ComputePowerFlowFromEH(ByRef dPower As Double)
    If g_n1 <= 0 Or g_n2 <= 0 Then
        dPower = 0#
        Exit Sub
    End If
    dPower = CComputePowerFlow( _
        CORE_AUTHOR, _
        g_iDir, g_n1, g_n2, _
        g_arrDw1(0), g_arrDw2(0), _
        g_aExRe(0), g_aExIm(0), _
        g_aEyRe(0), g_aEyIm(0), _
        g_aEzRe(0), g_aEzIm(0), _
        g_aHxRe(0), g_aHxIm(0), _
        g_aHyRe(0), g_aHyIm(0), _
        g_aHzRe(0), g_aHzIm(0))
End Sub

Private Function TempPath(ByVal sBase As String, ByVal iMode As Integer) As String
    TempPath = g_sTemp & "_" & sBase & "_" & CStr(iMode) & ".sig"
End Function

Private Sub CreateEmptyFile(ByVal sPath As String)
    On Error Resume Next
    Kill sPath
    On Error GoTo 0
    Dim f As Integer
    f = FreeFile
    Open sPath For Output As #f
    Close #f
End Sub

Private Sub BufferAppend(ByRef sBuf As String, ParamArray vals() As Variant)
    Dim i As Integer
    For i = 0 To UBound(vals)
        sBuf = sBuf & Format(CDbl(vals(i)), "0.000000e+000   ")
    Next i
    sBuf = sBuf & vbCrLf
End Sub

Private Sub BufferFlush(ByVal sPath As String, ByRef sBuf As String)
    If Len(sBuf) = 0 Then Exit Sub
    Dim f As Integer
    Err.Clear
    On Error Resume Next
    f = FreeFile
    Open sPath For Append As #f
    If Err <> 0 Then
        LogError "Failed to open the temporary result file (Err=" & Err.Number & "): " & sPath
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Print #f, sBuf;
    If Err <> 0 Then
        LogError "Failed to write the temporary result file (Err=" & Err.Number & "): " & sPath
        Close #f
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Close #f
    Err.Clear
    On Error GoTo 0
    sBuf = ""
End Sub

Private Sub ResetPhaseSweepState()
    Dim i As Long
    If g_nPhaseStateModes <= 0 Then Exit Sub
    For i = 1 To g_nPhaseStateModes
        g_abPhaseDead(i) = False
        g_anPhaseFlat(i) = 0
        g_abPhaseWarned(i) = False
    Next i
End Sub

Private Function RowBuffersPending() As Boolean
    Dim i As Long
    If g_nBufModes <= 0 Then Exit Function
    For i = 1 To g_nBufModes
        If Len(g_aBufBeta(i)) > 0 Or Len(g_aBufZpierce(i)) > 0 Or Len(g_aBufPhase(i)) > 0 Then
            RowBuffersPending = True
            Exit Function
        End If
    Next i
End Function

Private Sub ClearRowBuffers()
    Dim i As Long
    If g_nBufModes <= 0 Then Exit Sub
    For i = 1 To g_nBufModes
        g_aBufBeta(i) = ""
        g_aBufZpierce(i) = ""
        g_aBufPhase(i) = ""
    Next i
End Sub

Private Sub InitLogFile()
    LogInit
    If Len(g_sLogFile) = 0 Then Exit Sub

    If Dir(g_sLogFile) = "" Then CreateEmptyFile g_sLogFile

    Dim fTime As Integer
    fTime = FreeFile
    Open g_sTemp & "_runtime_start.txt" For Output As #fTime
    Print #fTime, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #fTime

    LogInfo " ----- New parameter sweep started, start time: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ----- "
    LogInfo " ----- Log file: " & g_sLogFile & " ----- "
    LogInfo " ----- Path budget: Temp path length = " & Len(g_sTemp) & " characters, file-name tag budget = " & g_nFileTagBudget & " characters ----- "

    If g_nFileTagBudget < 12 Then LogWarning "Project path too deep: the file-name tag budget is only " & g_nFileTagBudget & " characters; move the project to a shorter directory, otherwise the calculation data may fail to be written"

    CheckCoreLibrary
End Sub

Private Sub LogInit()
    If g_bLogInited Then Exit Sub
    g_bLogInited = True
    g_nLogFileMinLevel = llInfo
    g_sLogLastError = ""
End Sub

Private Sub RotateLogIfNeeded()
    Dim nSize As Long

    If Len(g_sLogFile) = 0 Then Exit Sub

    On Error Resume Next
    nSize = FileLen(g_sLogFile)
    On Error GoTo 0
    If nSize < LOG_ROTATE_BYTES Then Exit Sub

    On Error Resume Next
    Name g_sLogFile As g_sLogFile & "." & Format$(Now, "yyyymmdd_hhnnss") & ".bak"
    On Error GoTo 0
End Sub

Private Function WriteLog(ByVal sLevel As String, ByVal sMsg As String) As Boolean
    Dim f As Integer
    Dim nErr As Long
    Dim sErr As String
    Dim sLine As String

    If Len(g_sLogFile) = 0 Then Exit Function

    If Len(sMsg) > LOG_MSG_MAX_LEN Then sMsg = Left$(sMsg, LOG_MSG_MAX_LEN) & " ...(truncated)"
    sLine = "[" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "] [" & sLevel & "] " & sMsg

    RotateLogIfNeeded

    On Error GoTo Failed
    f = FreeFile
    Open g_sLogFile For Append As #f
    Print #f, sLine
    Close #f
    WriteLog = True
    Exit Function

Failed:
    nErr = Err.Number
    sErr = Err.Description
    On Error Resume Next
    Close #f
    On Error GoTo 0

    g_sLogLastError = "Log write failed (Err=" & nErr & "): " & sErr & " | file: " & g_sLogFile
    Debug.Print g_sLogLastError
End Function

Private Sub LogMessage(ByVal nLevel As LogLevel, ByVal sMsg As String, _
    Optional ByVal bAbortAfter As Boolean = False)
    Dim sTag As String
    Dim sPfx As String
    Dim bToWindow As Boolean

    If Len(Trim$(sMsg)) = 0 Then sMsg = "(empty message)"

    Select Case nLevel
        Case llCritical
            sTag = LOG_TAG_CRITICAL
            sPfx = PFX_MSG_CRITICAL
            bToWindow = True
        Case llError
            sTag = LOG_TAG_ERROR
            sPfx = PFX_MSG_ERROR
            bToWindow = True
        Case llWarning
            sTag = LOG_TAG_WARNING
            sPfx = PFX_MSG_WARNING
            bToWindow = True
        Case llInfo
            sTag = LOG_TAG_INFO
            sPfx = PFX_MSG_INFO
            bToWindow = True
        Case Else
            sTag = LOG_TAG_DEBUG
            bToWindow = False
    End Select

    If nLevel >= g_nLogFileMinLevel Then WriteLog sTag, sMsg

    If bToWindow Then
        On Error Resume Next
        Select Case nLevel
            Case llCritical
                ReportErrorToWindow sPfx & sMsg
            Case llError
                ReportErrorToWindow sPfx & sMsg
            Case llWarning
                ReportWarningToWindow sPfx & sMsg
            Case Else
                ReportInformationToWindow sPfx & sMsg
        End Select
        On Error GoTo 0
    End If

    If bAbortAfter Then
        LogCritical "A fatal error was found! The macro has aborted the parameter sweep! Check the parameter settings before restarting the sweep or running the macro again!"
        ReportError MSG_ABORT_MACRO & sMsg & vbCrLf
    End If
End Sub

Private Sub LogDebug(ByVal sMsg As String)
    LogMessage llDebug, sMsg
End Sub

Private Sub LogInfo(ByVal sMsg As String)
    LogMessage llInfo, sMsg
End Sub

Private Sub LogWarning(ByVal sMsg As String)
    LogMessage llWarning, sMsg
End Sub

Private Sub LogError(ByVal sMsg As String)
    LogMessage llError, sMsg
End Sub

Private Sub LogCritical(ByVal sMsg As String, Optional ByVal bAbortAfter As Boolean = False)
    LogMessage llCritical, sMsg, bAbortAfter
End Sub

Public Sub Main()
    ParameterSweepWatch 0
    ParameterSweepWatch 1
    ParameterSweepWatch 2
End Sub
