'==============================================================================
' 林夕宏代码 (CST-LinXi-Macro)  —  版本检查 (She.bas)
'
' Copyright (c) 2026 Limorazp
' SPDX-License-Identifier: MIT
'
' 本文件以 MIT 许可证开源发布, 完整条款见仓库根目录 LICENSE。
'
'==============================================================================

Option Explicit

' 核心计算库校验口令与文件名
Private Const CORE_AUTHOR   As String = " Lin Xi and She & Me "
Private Const DLL_FILE_NAME As String = "LinXi.dll"

Declare Function CCoreVersion Lib "LinXi.dll" (ByVal author As String) As Long

' 外部配置文件 (版本信息的唯一来源, 必须为 ANSI/GBK 编码)
Private Const CFG_FILE_NAME   As String = "LinXi.ini"
Private Const CFG_SECTION_VER As String = "VERSION"

' ---- DLL 状态码 ----
Private Const DLL_OK        As Integer = 0     ' 正常
Private Const DLL_NOT_FOUND As Integer = 1     ' 文件不存在
Private Const DLL_CALL_FAIL As Integer = 2     ' 调用失败
Private Const DLL_TOO_OLD   As Integer = 3     ' 版本低于最低要求

' ---- 全局状态 ----
Private g_sVersionString As String             ' 宏程序版本号
Private g_sVersionLabel  As String             ' 版本标签
Private g_nDllMinVer     As Long               ' DLL 最低版本要求
Private g_sCfgPath       As String             ' 实际命中的配置文件路径
Private g_sProblems      As String             ' 累积的问题清单 (空 = 一切正常)

'==============================================================================
' 宏入口: 读配置 -> 查 DLL 版本 -> 无问题则显示版本信息, 否则显示问题清单
'==============================================================================
Sub Main()
    Dim bCfgOK As Boolean
    Dim iDllStatus As Integer
    Dim nDllVer As Long
    Dim sDllNote As String

    g_sVersionString = ""
    g_sVersionLabel = ""
    g_nDllMinVer = 0
    g_sCfgPath = ""
    g_sProblems = ""
    nDllVer = 0
    sDllNote = ""

    bCfgOK = LoadVersionFromIni()
    iDllStatus = QueryDllVersion(nDllVer, sDllNote)

    If Len(g_sProblems) = 0 Then
        ShowVersionDialog nDllVer
    Else
        ShowProblemDialog bCfgOK, iDllStatus, nDllVer, sDllNote
    End If
End Sub

' 一切正常时显示的版本信息窗口
Private Sub ShowVersionDialog(ByVal nDllVer As Long)
    Dim sCstVer As String

    sCstVer = CstVersionText()

    Begin Dialog UserDialog 440, 204, "林夕宏代码程序版本信息"

        Text 108, 12, 290, 16, "CST 慢波结构用户监视器宏程序"
        Text 178, 32, 200, 14, "版  本  信  息"

        GroupBox 20, 56, 400, 88, ""

        Text  32,  70, 376, 14, "宏程序版本 : " & g_sVersionString & "  (" & g_sVersionLabel & ")"
        Text  32,  92, 376, 14, "核心计算库 : " & DLL_FILE_NAME & " v" & CStr(nDllVer) & "  (最低版本要求 v" & CStr(g_nDllMinVer) & ")"
        Text  32, 114, 376, 14, "CST 版本 : " & sCstVer

        OKButton 170, 158, 100, 42

    End Dialog

    Dim dlg As UserDialog
    Dialog dlg
End Sub

' 出现问题时显示的诊断窗口: 问题清单 + 已取得的信息 + 排查建议
Private Sub ShowProblemDialog(ByVal bCfgOK As Boolean, ByVal iDllStatus As Integer, _
    ByVal nDllVer As Long, ByVal sDllNote As String)

    Dim sMsg As String

    sMsg = "检测到以下问题:" & vbCrLf & vbCrLf & g_sProblems & vbCrLf & vbCrLf & _
        "已取得的信息:" & vbCrLf & _
        "  宏程序版本 : " & VersionText(bCfgOK) & vbCrLf & _
        "  核心计算库 : " & DllText(iDllStatus, nDllVer, sDllNote) & vbCrLf & _
        "  CST 版本   : " & CstVersionText() & vbCrLf & vbCrLf & _
        "提示: 重新运行安装脚本(双击我自动安装.bat)可恢复默认部署;" & vbCrLf & _
        "      若刚编辑过 LinXi.ini, 请确认它保存为 ANSI/GBK 编码。"

    MsgBox sMsg, vbCritical, "版本查询 - 检测到问题"
End Sub

' 宏程序版本文本; 配置未取到时返回"(未取得)"
Private Function VersionText(ByVal bOK As Boolean) As String
    If bOK Then
        VersionText = g_sVersionString & " (" & g_sVersionLabel & ")"
    Else
        VersionText = "(未取得)"
    End If
End Function

' 核心计算库的版本描述文本 (按状态区分五种情形)
Private Function DllText(ByVal iStatus As Integer, ByVal nVer As Long, ByVal sNote As String) As String
    If iStatus = DLL_OK Then
        DllText = DLL_FILE_NAME & " v" & CStr(nVer) & " (最低要求 v" & CStr(g_nDllMinVer) & ")"
    ElseIf nVer > 0 Then
        DllText = DLL_FILE_NAME & " v" & CStr(nVer) & " —— " & sNote
    ElseIf Len(sNote) > 0 Then
        DllText = DLL_FILE_NAME & " (" & sNote & ")"
    Else
        DllText = "(未取得)"
    End If
End Function

' 从 LinXi.ini 读取版本号、版本标签与 DLL 最低版本, 逐项校验后写入全局变量
Private Function LoadVersionFromIni() As Boolean
    Dim sVer As String
    Dim sLbl As String
    Dim sDllMin As String

    LoadVersionFromIni = False
    g_sCfgPath = CfgFind()
    sVer = ""
    sLbl = ""
    sDllMin = ""

    If g_sCfgPath = "" Then
        AddProblem "4 个候选位置均未找到外部配置文件 " & CFG_FILE_NAME & _
            ", 因而无法确定当前宏程序版本。" & vbCrLf & _
            "    宏按下列顺序查找该文件:" & vbCrLf & CfgCandidateList()
        Exit Function
    End If

    If Not CfgRead(g_sCfgPath, sVer, sLbl, sDllMin) Then
        AddProblem "外部配置文件无法读取(可能被占用或权限不足):" & vbCrLf & _
            "    " & g_sCfgPath
        Exit Function
    End If

    If Len(sVer) = 0 Or Len(sLbl) = 0 Or Len(sDllMin) = 0 Then
        AddProblem "配置文件中未读到 [Version] 节必填键 VersionString / VersionLabel / DllMinVersion:" & vbCrLf & _
            "    " & g_sCfgPath & vbCrLf & _
            "    若该文件刚被编辑过, 请确认保存为 ANSI/GBK 编码 —— 另存为 UTF-8 会读不到任何键。"
        Exit Function
    End If

    If Not IsNumeric(sDllMin) Then
        AddProblem "配置文件中 DllMinVersion 不是有效的版本号: """ & sDllMin & _
            """ (应为正整数, 例如 122)。"
        Exit Function
    End If

    g_nDllMinVer = CLng(Val(sDllMin))
    If g_nDllMinVer <= 0 Then
        AddProblem "配置文件中 DllMinVersion 必须为正整数, 当前为 """ & sDllMin & """。"
        Exit Function
    End If

    g_sVersionString = sVer
    g_sVersionLabel = sLbl
    LoadVersionFromIni = True
End Function

' 追加一条问题到问题清单 g_sProblems
Private Sub AddProblem(ByVal sMsg As String)
    If Len(g_sProblems) > 0 Then g_sProblems = g_sProblems & vbCrLf & vbCrLf
    g_sProblems = g_sProblems & "· " & sMsg
End Sub

' LinXi.ini 的第 iIndex 个候选路径 (安装目录 / 宏目录 / 工程目录)
Private Function CfgCandidatePath(ByVal iIndex As Integer) As String
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

    CfgCandidatePath = sPath
End Function

' 候选路径清单文本, 用于报错时提示宏都去哪里找过
Private Function CfgCandidateList() As String
    Dim sList As String
    Dim i As Integer

    sList = ""
    For i = 1 To 4
        If Len(CfgCandidatePath(i)) > 0 Then
            sList = sList & "    " & i & ". " & CfgCandidatePath(i) & vbCrLf
        End If
    Next i
    If Len(sList) > Len(vbCrLf) Then sList = Left$(sList, Len(sList) - Len(vbCrLf))

    CfgCandidateList = sList
End Function

' 文件是否存在 (Dir 失败一律按不存在处理)
Private Function DiskFileExists(ByVal sPath As String) As Boolean
    Dim sHit As String

    DiskFileExists = False
    sHit = ""
    Err.Clear
    On Error Resume Next
    sHit = Dir(sPath)
    If Err.Number = 0 And Len(sHit) > 0 Then DiskFileExists = True
    Err.Clear
    On Error GoTo 0
End Function

' 按候选顺序定位 LinXi.ini
Private Function CfgFind() As String
    Dim i As Integer
    Dim sPath As String

    CfgFind = ""
    For i = 1 To 4
        sPath = CfgCandidatePath(i)
        If Len(sPath) > 0 Then
            If DiskFileExists(sPath) Then
                CfgFind = sPath
                Exit Function
            End If
        End If
    Next i
End Function

' 逐行解析 ini, 取 [Version] 节的 VersionString / VersionLabel / DllMinVersion
' 注意: 文件必须为 ANSI/GBK 编码, UTF-8 存档会读不到任何键
Private Function CfgRead(ByVal sPath As String, ByRef sVer As String, _
    ByRef sLbl As String, ByRef sDllMin As String) As Boolean

    Dim f As Integer
    Dim sLine As String
    Dim sSection As String
    Dim sKey As String
    Dim sVal As String
    Dim nPos As Long
    Dim bOpened As Boolean

    CfgRead = False
    sVer = ""
    sLbl = ""
    sDllMin = ""
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
                    If sSection = CFG_SECTION_VER And Len(sVal) > 0 Then
                        Select Case sKey
                            Case "VERSIONSTRING"
                                sVer = sVal
                            Case "VERSIONLABEL"
                                sLbl = sVal
                            Case "DLLMINVERSION"
                                sDllMin = sVal
                        End Select
                    End If
                End If
            End If
        End If
    Loop
    Err.Clear
    Close #f
    On Error GoTo 0

    CfgRead = True
End Function

' 校验核心计算库: 定位路径、确认文件存在、调用 CCoreVersion 并比对最低版本
' 返回 DLL_OK / DLL_NOT_FOUND / DLL_CALL_FAIL / DLL_TOO_OLD, 失败原因写入 sNote
Private Function QueryDllVersion(ByRef nVer As Long, ByRef sNote As String) As Integer
    Dim sDllPath As String
    Dim nErrNum As Long
    Dim nDllErr As Long
    Dim sErrDesc As String

    nVer = 0
    sNote = ""
    QueryDllVersion = DLL_OK

    sDllPath = DllFilePath()
    If Len(sDllPath) = 0 Then
        QueryDllVersion = DLL_NOT_FOUND
        sNote = "无法定位文件路径"
        AddProblem "无法定位核心计算库 " & DLL_FILE_NAME & " 的路径: GetInstallPath 调用失败。"
        Exit Function
    End If

    If Not DiskFileExists(sDllPath) Then
        QueryDllVersion = DLL_NOT_FOUND
        sNote = "文件不存在"
        AddProblem "未找到核心计算库文件:" & vbCrLf & _
            "    " & sDllPath & vbCrLf & _
            "    请重新运行安装脚本(双击我自动安装.bat)完成部署。"
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    nVer = CCoreVersion(CORE_AUTHOR)
    nErrNum = Err.Number
    nDllErr = Err.LastDLLError
    sErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    If nErrNum <> 0 Or nVer <= 0 Then
        QueryDllVersion = DLL_CALL_FAIL
        sNote = "调用失败"
        AddProblem "核心计算库文件存在, 但 CCoreVersion 调用失败:" & vbCrLf & _
            "    Err.Number   = " & CStr(nErrNum) & vbCrLf & _
            "    LastDLLError = " & CStr(nDllErr) & vbCrLf & _
            "    Description  = " & sErrDesc & vbCrLf & _
            "    常见原因: DLL 为 32 位(应为 64 位) / 导出函数名不匹配 / 依赖库缺失。"
        Exit Function
    End If

    If nVer < g_nDllMinVer Then
        QueryDllVersion = DLL_TOO_OLD
        sNote = "版本过旧"
        AddProblem "核心计算库版本过旧: 当前 v" & CStr(nVer) & _
            ", 本程序要求不低于 v" & CStr(g_nDllMinVer) & vbCrLf & _
            "    (该要求写在 " & CFG_FILE_NAME & " 的 DllMinVersion)"
        Exit Function
    End If
End Function

' 核心计算库的预期路径: <CST安装目录>\AMD64\LinXi.dll
Private Function DllFilePath() As String
    Dim sPath As String

    sPath = ""
    Err.Clear
    On Error Resume Next
    sPath = GetInstallPath & "\AMD64\" & DLL_FILE_NAME
    If Err.Number <> 0 Then
        sPath = ""
        Err.Clear
    End If
    On Error GoTo 0

    DllFilePath = sPath
End Function

' CST 版本号文本, 取不到时返回"(无法获取)"
Private Function CstVersionText() As String
    Dim sVer As String

    sVer = ""
    Err.Clear
    On Error Resume Next
    sVer = GetApplicationVersion()
    If Err.Number <> 0 Then sVer = ""
    Err.Clear
    On Error GoTo 0

    sVer = Trim$(sVer)
    If Len(sVer) = 0 Then sVer = "(无法获取)"
    CstVersionText = sVer
End Function
