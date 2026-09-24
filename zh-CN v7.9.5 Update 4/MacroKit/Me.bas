'==============================================================================
' 林夕宏代码 (CST-LinXi-Macro)  —  慢波结构用户监视器配置向导 (Me.bas)
'
' Copyright (c) 2026 Limorazp
' SPDX-License-Identifier: MIT
'
' 本文件以 MIT 许可证开源发布, 完整条款见仓库根目录 LICENSE。
'
'==============================================================================

Option Explicit

' 监视器宏的部署位置: 与 LinXi.dll / LinXi.ini 同在 CST 安装目录的 AMD64 下
Private Const WATCH_FILE_NAME As String = "LinXi_Watch.bas"

' 监视器宏旧版(<= v7.9.5.1)的部署位置, 仅作兼容回退
Private Const WATCH_LEGACY_DIR As String = "\Library\Macros\Solver\E-Solver\"
Private Const WATCH_LEGACY_NAME As String = "Model.pfc"

' 监视器宏复制进工程后的固定文件名, CST 要求为 Model.pfc
Private Const PROJECT_MACRO_NAME As String = "Model.pfc"

Private Const CFG_FILE_NAME   As String = "LinXi.ini"
Private Const CFG_SECTION_VER As String = "VERSION"

' ---- 版本信息与配置状态 (全部来自 LinXi.ini) ----
Private g_sVersionString As String
Private g_sVersionLabel  As String
Private g_sAuthorName    As String
Private g_sCfgPath       As String
Private g_sCfgNote       As String

'==============================================================================
' 宏入口
'   1. 读 LinXi.ini 取得版本信息 (取不到直接中止, 不做半成品配置)
'   2. 定位监视器宏源文件
'   3. 向用户确认 (会清空现有结果)
'   4. 注册监视器 + 把监视器宏复制为工程的 Model.pfc
'==============================================================================
Sub Main()
    Dim bWatchExists As Boolean
    Dim bRegistered  As Boolean
    Dim sWatchMacro  As String

    ' --- 第一步: 版本信息必须来自外部配置, 否则中止 ---
    If Not LoadCfgForWizard() Then Exit Sub

    ' --- 第二步: 先确认监视器宏源文件存在, 避免用户确认完才发现部署不全 ---
    sWatchMacro = WatchMacroFind()
    If Len(sWatchMacro) = 0 Then
        ShowWatchMissingDialog
        Exit Sub
    End If

    ' --- 第三步: 向用户确认 (添加监视器会清空工程现有结果) ---
    If Not ShowWelcomeDialog() Then
        MsgBox "用户取消添加慢波结构用户监视器。", vbInformation, "已取消"
        Exit Sub
    End If

    ' --- 第四步: 注册监视器并部署宏文件 ---
    bWatchExists = CopyWatchMacro(sWatchMacro)

    If Not bWatchExists Then
        MsgBox "监视器宏部署失败, 未能写入工程目录:" & vbCrLf & vbCrLf & _
            ProjectMacroPath() & vbCrLf & vbCrLf & _
            "请确认工程目录可写, 或重新运行安装脚本(双击我自动安装.bat)。", _
            vbCritical, "添加失败"
        Exit Sub
    End If

    bRegistered = AddWatchToHistory()

    If bRegistered Then
        MsgBox "您已成功添加慢波结构用户监视器!" & vbCrLf & vbCrLf & _
            "监视器宏: " & sWatchMacro & vbCrLf & _
            "工程内文件: " & ProjectMacroPath() & vbCrLf & vbCrLf & _
            "直接开启扫描参数任务即可, 无需修改宏代码!", vbInformation, "完成"
    Else
        MsgBox "监视器宏已复制到工程, 但写入项目历史记录失败。" & vbCrLf & vbCrLf & _
            "请重新运行本宏(Define LinXi Macro)重试。", vbExclamation, "部分完成"
    End If
End Sub

' 把监视器注册进项目历史记录; 失败不抛错, 由调用方提示
Private Function AddWatchToHistory() As Boolean
    AddWatchToHistory = False

    On Error Resume Next
    AddToHistory "Add Watch: LinXi Macro", "ParameterSweep.AddUserdefinedWatch"
    If Err.Number = 0 Then AddWatchToHistory = True
    Err.Clear
    On Error GoTo 0
End Function

' 欢迎/确认对话框: 说明功能与注意事项, 并提示"添加监视器会清空现有结果"
Private Function ShowWelcomeDialog() As Boolean

    Begin Dialog UserDialog 640, 395, "慢波结构用户监视器配置向导"

        Text 180, 20, 300, 14, "欢迎使用慢波结构用户监视器配置向导!"

        GroupBox 20, 42, 600, 71, "[ 功能说明 ]"

        Text 40, 63, 560, 14, "该监视器可自动计算慢波结构的相速度与耦合阻抗等后处理结果。"
        Text 40, 91, 560, 14, "最终生成三类曲线: 布里渊图、归一化相速度曲线和耦合阻抗曲线。"

        GroupBox 20, 121, 600, 127, "[ 温馨提示 ]"

        Text 40, 142, 560, 14, "1. 在一个工程文件中只需配置一次监视器即可。"
        Text 40, 170, 560, 14, "2. 删除历史树中的监视器步骤，不等于移除宏程序本身。"
        Text 40, 198, 560, 14, "3. 使用前请确认已将 LinXi_Watch.bas 与 LinXi.dll 部署到 CST 的 AMD64 目录。"
        Text 40, 226, 560, 14, "4. 扫参期间请勿操作 CST 窗口，建议单开窗口、单任务运行。"

        GroupBox 20, 256, 600, 43, "[ 重要警告 ]"

        Text 40, 277, 560, 14, "添加监视器后，现有的数据结果将被删除，是否继续？"

        CancelButton 200, 320, 100, 42
        OKButton     340, 320, 100, 42

        Text  20, 372, 300, 14, "版本: " & g_sVersionString & " (" & g_sVersionLabel & ")"
        Text 512, 372, 200, 14, "编写者:  " & g_sAuthorName

    End Dialog

    Dim dlg As UserDialog
    If Dialog(dlg) Then
        ShowWelcomeDialog = True
    Else
        ShowWelcomeDialog = False
    End If
End Function

' 工程目录下监视器宏的落点 (CST 要求文件名为 Model.pfc)
Private Function ProjectMacroPath() As String
    Dim sProj As String
    Dim sPath As String

    sPath = ""
    Err.Clear
    On Error Resume Next
    sProj = GetProjectPath("Model3D")
    If Err.Number <> 0 Then sProj = ""
    Err.Clear
    On Error GoTo 0

    If Len(sProj) = 0 Then
        ProjectMacroPath = ""
        Exit Function
    End If

    If Right$(sProj, 1) <> "\" Then sProj = sProj & "\"
    ProjectMacroPath = sProj & PROJECT_MACRO_NAME
End Function

' 定位监视器宏源文件: 优先 AMD64 (与 LinXi.dll 同目录), 再回退旧版部署位置
Private Function WatchMacroFind() As String
    Dim i As Integer
    Dim sPath As String

    WatchMacroFind = ""
    For i = 1 To 2
        sPath = WatchMacroCandidate(i)
        If Len(sPath) > 0 Then
            If DiskFileExists(sPath) Then
                WatchMacroFind = sPath
                Exit Function
            End If
        End If
    Next i
End Function

' 监视器宏候选路径: 1 = AMD64 (当前版本), 2 = E-Solver 宏目录 (旧版兼容)
Private Function WatchMacroCandidate(ByVal iIndex As Integer) As String
    Dim sPath As String

    sPath = ""
    Err.Clear
    On Error Resume Next
    Select Case iIndex
        Case 1
            sPath = GetInstallPath & "\AMD64\" & WATCH_FILE_NAME
        Case 2
            sPath = GetInstallPath & WATCH_LEGACY_DIR & WATCH_LEGACY_NAME
    End Select
    If Err.Number <> 0 Then sPath = ""
    Err.Clear
    On Error GoTo 0

    WatchMacroCandidate = sPath
End Function

' 监视器宏缺失时的提示: 列出所有候选路径与排查建议
Private Sub ShowWatchMissingDialog()
    Dim sMsg As String
    Dim i As Integer

    sMsg = "错误: 未找到监视器宏 " & WATCH_FILE_NAME & ", 添加监视器的操作已中止!" & vbCrLf & vbCrLf & _
        "宏按下列顺序查找该文件, 请确认至少有一处存在:" & vbCrLf
    For i = 1 To 2
        sMsg = sMsg & "  " & i & ". " & WatchMacroCandidate(i) & vbCrLf
    Next i
    sMsg = sMsg & vbCrLf & _
        "提示: 请重新运行安装脚本(双击我自动安装.bat), 它会将 " & WATCH_FILE_NAME & _
        " 与 " & CFG_FILE_NAME & " 一起部署到 CST 安装目录的 AMD64 下。"

    MsgBox sMsg, vbCritical, "未找到监视器宏"
End Sub

' 把监视器宏复制为工程的 Model.pfc; 返回是否成功
Private Function CopyWatchMacro(ByVal sSrc As String) As Boolean
    Dim sDst As String

    CopyWatchMacro = False

    sDst = ProjectMacroPath()
    If Len(sDst) = 0 Then Exit Function
    If Not DiskFileExists(sSrc) Then Exit Function

    Err.Clear
    On Error Resume Next
    FileCopy sSrc, sDst
    If Err.Number = 0 Then CopyWatchMacro = True
    Err.Clear
    On Error GoTo 0
End Function

' 文件是否真实存在 (Dir 失败一律按不存在处理)
Private Function DiskFileExists(ByVal sPath As String) As Boolean
    Dim sHit As String

    DiskFileExists = False
    If Len(sPath) = 0 Then Exit Function

    sHit = ""
    Err.Clear
    On Error Resume Next
    sHit = Dir(sPath)
    If Err.Number = 0 And Len(sHit) > 0 Then DiskFileExists = True
    Err.Clear
    On Error GoTo 0
End Function

' 为向导读取 LinXi.ini 的版本与作者信息; 取不到则弹窗并中止本宏
Private Function LoadCfgForWizard() As Boolean
    Dim sVer As String
    Dim sLbl As String
    Dim sAuthor As String

    LoadCfgForWizard = False
    g_sVersionString = ""
    g_sVersionLabel = ""
    g_sAuthorName = ""
    g_sCfgNote = ""
    g_sCfgPath = CfgFind()
    sVer = ""
    sLbl = ""
    sAuthor = ""

    If g_sCfgPath = "" Then
        g_sCfgNote = "4 个候选位置均未找到外部配置文件 " & CFG_FILE_NAME
    ElseIf Not CfgRead(g_sCfgPath, sVer, sLbl, sAuthor) Then
        g_sCfgNote = "外部配置文件无法读取(可能被占用或权限不足): " & g_sCfgPath
    ElseIf Len(sVer) = 0 Or Len(sLbl) = 0 Or Len(sAuthor) = 0 Then
        g_sCfgNote = "配置文件中未读到 [Version] 节必填键 VersionString / VersionLabel / Author: " & g_sCfgPath
    Else
        g_sVersionString = sVer
        g_sVersionLabel = sLbl
        g_sAuthorName = sAuthor
    End If

    If Len(g_sVersionString) = 0 Then
        MsgBox CfgErrorText(), vbCritical, "宏程序已中止"
        Exit Function
    End If

    LoadCfgForWizard = True
End Function

' 配置读取失败的提示文本 (含 4 个候选路径与编码要求)
Private Function CfgErrorText() As String
    Dim sMsg As String
    Dim i As Integer

    sMsg = "错误: 版本信息必须来自外部配置文件 " & CFG_FILE_NAME & _
        ", 但当前无法取得, 添加监视器的操作已中止!" & vbCrLf & vbCrLf & _
        "原因: " & g_sCfgNote & vbCrLf & vbCrLf & _
        "宏按下列顺序查找该文件, 请确认至少有一处存在且可读:" & vbCrLf
    For i = 1 To 4
        If Len(CfgCandidatePath(i)) > 0 Then
            sMsg = sMsg & "  " & i & ". " & CfgCandidatePath(i) & vbCrLf
        End If
    Next i
    sMsg = sMsg & vbCrLf & _
        "提示: 该文件必须保存为 ANSI/GBK 编码; 若被另存为 UTF-8, 将读不到任何键。" & vbCrLf & _
        "      请重新运行安装脚本(双击我自动安装.bat)完成部署。"

    CfgErrorText = sMsg
End Function

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

' 文件是否存在
Private Function CfgExists(ByVal sPath As String) As Boolean
    CfgExists = DiskFileExists(sPath)
End Function

' 按候选顺序定位 LinXi.ini
Private Function CfgFind() As String
    Dim i As Integer
    Dim sPath As String

    CfgFind = ""
    For i = 1 To 4
        sPath = CfgCandidatePath(i)
        If Len(sPath) > 0 Then
            If CfgExists(sPath) Then
                CfgFind = sPath
                Exit Function
            End If
        End If
    Next i
End Function

' 逐行解析 ini 的 [Version] 节, 取出版本号、版本标签与作者
' 注意: 文件必须为 ANSI/GBK 编码, UTF-8 存档会读不到任何键
Private Function CfgRead(ByVal sPath As String, ByRef sVer As String, _
    ByRef sLbl As String, ByRef sAuthor As String) As Boolean

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
    sAuthor = ""
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
                ' 整行注释, 忽略
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
                            Case "AUTHOR"
                                sAuthor = sVal
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
