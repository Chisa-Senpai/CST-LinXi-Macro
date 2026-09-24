'==============================================================================
' 林夕宏代码 (CST-LinXi-Macro)  —  CST 本征模慢波结构用户监视器宏
'
' Copyright (c) 2026 Limorazp
' SPDX-License-Identifier: MIT
'
' 本文件以 MIT 许可证开源发布。在遵守许可证的前提下, 可自由使用、复制、修改、
' 合并、发布、分发、再许可和/或销售本软件的副本。上述版权声明与本许可声明
' 应包含在本软件的所有副本或实质性部分中。完整条款见仓库根目录 LICENSE。
'
'==============================================================================

Option Explicit

' ---- 核心计算库校验口令: 传给 LinXi.dll 每个导出函数, 用于校验调用方 ----
Private Const CORE_AUTHOR As String = " Lin Xi and She & Me "

' ---- 物理常量与单位换算 ----
Private Const CLight   As Double = 299792458#      ' 真空光速 (m/s)
Private Const DegToRad As Double = Pi / 180#       ' 角度转弧度

' ---- CST 界面等待延时 (秒); 机器较慢导致畸形曲线时优先调大此值 ----
Private Const WAIT_SEC As Double = 0.1

' ---- 容量与重试上限 ----
Private Const MAX_MODE_PARAMS   As Integer = 1000      ' Macro_Mode 参数登记上限
Private Const MAX_PTS_PER_DIR   As Long = 1000         ' 采样网格单方向最大点数
Private Const MAX_CROSS_POINTS  As Long = 200000       ' 横截面采样点总数上限
Private Const MAX_FREQ_RETRY    As Integer = 30        ' 读场频率重试次数
Private Const FREQ_RETRY_WAIT   As Double = 0.2        ' 每次重试前的等待 (秒)

' ---- 相位扫描有效性判据 ----
Private Const FREQ_FLAT_TOL     As Double = 1E-09      ' 判定频率"未变化"的相对容差
Private Const FREQ_FLAT_MAX     As Integer = 5         ' 连续未变化点数达到该值即判扫描无效
Private Const KC_SANITY_MAX     As Double = 1E+08      ' 耦合阻抗合理上限 (Ohm), 超过记为无效

' ---- 路径长度预算: 规避 Windows 260 字符限制 ----
Private Const MAX_PATH_SAFE        As Integer = 250    ' Temp 路径 + 文件名允许的总长
Private Const FILE_TAG_FIXED       As Integer = 24     ' 文件名中固定部分占用的字符数
Private Const MIN_FILE_TAG_BUDGET  As Integer = 18     ' 标签可用字符数下限, 低于该值无法生成安全文件名

' ---- 日志等级; 数值越大越严重 ----
Private Enum LogLevel
    llDebug = 0
    llInfo = 1
    llWarning = 2
    llError = 3
    llCritical = 4
End Enum

' ---- 日志文件中的等级标签 (定宽, 便于对齐阅读) ----
Private Const LOG_TAG_DEBUG        As String = " DEBUG "
Private Const LOG_TAG_INFO         As String = " INFO  "
Private Const LOG_TAG_WARNING      As String = " WARN  "
Private Const LOG_TAG_ERROR        As String = " ERROR "
Private Const LOG_TAG_CRITICAL     As String = " CRIT! "

' ---- 输出到 CST 消息窗口的前缀 ----
Private Const PFX_MSG_INFO         As String = "MacroMsg [INFO] "
Private Const PFX_MSG_WARNING      As String = "MacroMsg [WARN] "
Private Const PFX_MSG_ERROR        As String = "MacroMsg [ERROR] "
Private Const PFX_MSG_CRITICAL     As String = "MacroMsg [CRIT] "

' ---- 日志文件策略: 位于工程 Temp 目录, 超过 5 MB 自动滚动备份 ----
Private Const LOG_NAME             As String = "_macro_log.txt"
Private Const LOG_MSG_MAX_LEN      As Long = 10000     ' 单条日志最大长度, 超出截断
Private Const LOG_ROTATE_BYTES     As Long = 5242880   ' 滚动阈值 (5 MB)
Private Const MSG_ABORT_MACRO      As String = vbCrLf & vbCrLf & "发现致命错误! 宏程序已中止扫参程序进行!" & vbCrLf & _
                                               "请检查各项参数设置后再重新开始扫参程序或执行宏代码程序!" & vbCrLf & _
                                               "错误信息: "

' ---- 结果曲线在临时文件名与结果树中使用的名称 ----
Private Const F_BETA    As String = "beta"                       ' 相位常数 beta
Private Const F_ZPIERCE As String = "ZpierceAvg"                 ' 皮尔斯耦合阻抗
Private Const F_VPHASE  As String = "vphase"                     ' 归一化相速度
Private Const F_PHASE   As String = "phase"                      ' 扫描相位
Private Const F_RESULT_GROUP   As String = "PlotOutput LinXi Macro Result"  ' 结果树分组名

' ---- 外部配置文件 (版本信息的唯一来源) ----
Private Const CFG_FILE_NAME    As String = "LinXi.ini"
Private Const CFG_SECTION_VER  As String = "VERSION"

' ---- 功率流计算方式 ----
Private Const PW_SRC_FROM_EH  As Integer = 1     ' 使能位整数位编码: 间接法 (由 E、H 场先插值后叉积再积分)
Private Const PW_SRC_FROM_CST As Integer = 2     ' 使能位整数位编码: 原生法 (用 CST 生成的功率流场)
Private Const PW_SRC_INVALID  As Integer = 0     ' 禁用 / 非法 (保留: 使能位不得取 0)
Private Const PW_SRC_INDIRECT As Integer = 0     ' 内部编码: 间接法, 等于 PW_SRC_FROM_EH - 1
Private Const PW_SRC_NATIVE   As Integer = 1     ' 内部编码: 原生法, 等于 PW_SRC_FROM_CST - 1

' ---- 耦合阻抗计算区域 (使能位 Macro_SweepWatch_Enable 的一位小数位) ----
Private Const REGION_BOTH     As Integer = 0     ' 行波区 + 返波区: 不判断频率升降
Private Const REGION_FORWARD  As Integer = 1     ' 仅行波区: 要求频率随 phase 上升
Private Const REGION_BACKWARD As Integer = 2     ' 仅返波区: 要求频率随 phase 下降

' ---- 使能位取值解析: 整数位 + 一位小数位, 其余取值一律判为非法 ----
Private Const ENABLE_FLAG_SCALE  As Long = 10    ' 使能位放大倍数, 用于取出整数位与小数位
Private Const ENABLE_FLAG_EPS    As Double = 1E-06   ' 小数位取整容差

'==============================================================================
' 全局状态
'==============================================================================

' ---- 模式选择与几何状态 ----
Private g_bIsTetra      As Boolean             ' True = 四面体网格, False = 六面体
Private g_bFirstFieldLog As Boolean            ' 首次成功解析场文件的标记
Private g_bAllModes     As Boolean             ' True = 计算全部模式
Private g_abModeFlag()  As Boolean             ' 各模式是否参与计算
Private g_bGridFailed   As Boolean             ' 采样网格构建失败标记
Private g_abPhaseDead() As Boolean             ' 各模式的相位扫描是否已判为无效
Private g_anPhaseFlat() As Long                ' 各模式频率连续未变化的点数
Private g_abPhaseFlatWarned() As Boolean       ' 各模式是否已给出过"频率未随 phase 变化"告警
Private g_abKcWarned() As Boolean              ' 各模式是否已给出过"Kc 超过合理上限"告警
Private g_nPhaseStateModes As Long             ' 相位状态数组当前容量

' ---- 路径与结果分组 ----
Private g_sPointsFile   As String              ' 采样点坐标文件 (_field_points.txt)
Private g_sLastGroupKey As String              ' 上一个参数组合键, 用于检测组合切换
Private g_sUnit         As String              ' 几何单位名称 (mm / um ...)
Private g_sDirLabel    As String               ' 周期方向标签 X / Y / Z
Private g_sGroupPath   As String               ' 参数组合在结果树中的路径
Private g_sGroupKey    As String               ' 参数组合键 (phase 之外的扫描变量)
Private g_sLogFile      As String              ' 日志文件完整路径
Private g_sTemp         As String              ' 工程 Temp 目录
Private g_sResult       As String              ' 工程 Result 目录
Private g_sFileTag      As String              ' 本次组合的文件名标签
Private g_nFileTagBudget As Long               ' 文件名标签可用字符数预算

' ---- 日志运行状态 ----
Private g_bLogInited        As Boolean
Private g_nLogFileMinLevel  As LogLevel        ' 写盘的最低等级, 低于该等级只进调试输出
Private g_sLogLastError     As String          ' 最近一次写盘失败信息, 结束时汇总提示

' ---- 版本信息 ----
Private g_sVersionString As String
Private g_sVersionLabel  As String
Private g_sAuthorName    As String
Private g_sReleaseDate   As String
Private g_sPlatform      As String
Private g_sCstVerMin     As String
Private g_sCstVerMax     As String
Private g_sCfgPath       As String             ' 实际命中的配置文件路径
Private g_sCfgNote       As String             ' 配置读取失败原因
Private g_sCfgVerRaw     As String
Private g_sCfgLblRaw     As String
Private g_sCfgDllMinRaw  As String
Private g_nDllMinVer     As Long               ' DLL 最低版本要求
Private g_bCfgOK         As Boolean            ' 配置是否全部读取成功

' ---- 计算选项 ----
Private g_iDir              As Integer         ' 周期方向: 1=X 2=Y 3=Z
Private g_iNumModes         As Integer         ' 求解器请求的本征模数量
Private g_aSelectedModes()  As Integer         ' 用户选定的模式编号
Private g_nSelectedModes    As Integer
Private g_iPowerFlowSrc     As Integer         ' 功率流计算方式 (0 = 间接法, 1 = 原生法)
Private g_iRegionMode       As Integer         ' 计算区域 (0 = 行波区 + 返波区, 1 = 行波区, 2 = 返波区)

' ---- 场数据缓冲: 按采样点顺序存放复矢量, Re/Im 分离 ----
Private g_aExRe() As Double, g_aExIm() As Double
Private g_aEyRe() As Double, g_aEyIm() As Double
Private g_aEzRe() As Double, g_aEzIm() As Double
Private g_aHxRe() As Double, g_aHxIm() As Double
Private g_aHyRe() As Double, g_aHyIm() As Double
Private g_aHzRe() As Double, g_aHzIm() As Double
Private g_aPxRe() As Double, g_aPxIm() As Double
Private g_aPyRe() As Double, g_aPyIm() As Double
Private g_aPzRe() As Double, g_aPzIm() As Double

' ---- 临时结果行缓冲: 按模式累积, 攒够一批再落盘以减少磁盘 IO ----
Private g_aBufBeta()    As String
Private g_aBufZpierce() As String
Private g_aBufPhase()   As String
Private g_nBufModes     As Long

' ---- 采样网格几何: 各方向网格中点与权重 ----
Private g_arrDw1() As Double, g_arrDw2() As Double
Private g_arrMidL() As Double, g_arrDwL() As Double
Private g_cMin(1 To 3) As Double, g_cMax(1 To 3) As Double   ' 计算域包围盒
Private g_xRef As Double, g_yRef As Double, g_zRef As Double ' 耦合阻抗参考位置
Private g_dPitchSI   As Double                 ' 周期长度 (m)
Private g_dPitchUnit As Double                 ' 周期长度 (几何单位)
Private g_dUnitToSI  As Double                 ' 几何单位到米的换算系数

' ---- 采样网格规模 ----
Private g_n1 As Long        ' 横向方向 1 网格数
Private g_n2 As Long        ' 横向方向 2 网格数
Private g_nCross As Long    ' 横截面采样点总数 = n1 * n2
Private g_nL As Long        ' 周期方向网格数

'==============================================================================
' 核心计算库 LinXi.dll 导出函数声明
' 首参 CORE_AUTHOR 为校验口令; 返回 0 = 正常, 非 0 = 错误码。
' 注: 此处刻意不加 Public/Private 修饰, 与原分发版本保持一致。
'==============================================================================

Declare Function CCoreVersion Lib "LinXi.dll" (ByVal author As String) As Long

' 解析坐标点采样得到的矢量场文件 (按点序返回 Ex/Ey/Ez 的实部与虚部)
Declare Function CParseFieldFile Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal sPath As String, ByVal nExpected As Long, _
    ByRef firstX As Double, ByRef firstY As Double, ByRef firstZ As Double, _
    ByRef exRe As Double, ByRef exIm As Double, _
    ByRef eyRe As Double, ByRef eyIm As Double, _
    ByRef ezRe As Double, ByRef ezIm As Double) As Long

' 解析 CST 原生功率流场文件 (按点序返回 Px/Py/Pz 的实部与虚部)
Declare Function CParsePowerFlowFile Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal sPath As String, ByVal nExpected As Long, _
    ByRef pxRe As Double, ByRef pxIm As Double, _
    ByRef pyRe As Double, ByRef pyIm As Double, _
    ByRef pzRe As Double, ByRef pzIm As Double) As Long

' 间接法: 由采样点上的 E、H 场求复坡印廷矢量并沿横截面加权积分, 返回净功率流
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

' 沿周期方向对场加权积分得到等效电压 V; 返回 0 表示成功
Declare Function CIntegrateLongitudinal Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal iDir As Long, ByVal nL As Long, ByVal nCross As Long, _
    ByVal beta As Double, ByVal unitToSI As Double, ByVal pitchSI As Double, _
    ByRef midL As Double, ByRef dwL As Double, _
    ByRef exRe As Double, ByRef exIm As Double, _
    ByRef eyRe As Double, ByRef eyIm As Double, _
    ByRef ezRe As Double, ByRef ezIm As Double, _
    ByRef vre As Double, ByRef vim As Double, ByRef eabs As Double) As Long

' 原生法: 对 CST 生成的功率流场沿横截面加权积分, 返回净功率流
Declare Function CIntegratePowerFlow Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal iDir As Long, ByVal n1 As Long, ByVal n2 As Long, _
    ByRef dw1 As Double, ByRef dw2 As Double, _
    ByRef pxRe As Double, ByRef pxIm As Double, _
    ByRef pyRe As Double, ByRef pyIm As Double, _
    ByRef pzRe As Double, ByRef pzIm As Double) As Double

' 四面体网格: 按最小/最大边长自动确定采样步长 h 与各方向点数
Declare Function CComputeTetraStep Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal len1 As Double, ByVal len2 As Double, ByVal lenL As Double, _
    ByVal minEdge As Double, ByVal maxEdge As Double, _
    ByVal maxCrossPoints As Long, ByVal maxPtsPerDir As Long, _
    ByRef n1 As Long, ByRef n2 As Long, ByRef nL As Long, _
    ByRef h As Double) As Long

' 四面体网格: 由计算域范围生成均匀采样网格与积分权重, 并写出采样点文件
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

' 六面体网格: 直接沿用 CST 网格坐标生成采样网格与积分权重
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

' 扫描结束后由临时数据生成四条曲线 (beta / phase / 相速度 / 耦合阻抗) 的曲线文件
Declare Function CBuildCurveFiles Lib "LinXi.dll" ( _
    ByVal author As String, _
    ByVal sTempPrefix As String, ByVal sSuffix As String, _
    ByVal iMode As Long, ByVal cLight As Double, _
    ByVal sOutBeta As String, ByVal sOutPhase As String, _
    ByVal sOutVp As String, ByVal sOutKc As String, _
    ByRef nBetaPt As Long, ByRef nPhasePt As Long, _
    ByRef nVpPt As Long, ByRef nKcPt As Long) As Long

'==============================================================================
' 宏入口
'==============================================================================

' CST 参数扫描监视器回调: 0=初始化 1=每个扫描点 2=扫描结束
Private Sub ParameterSweepWatch(ByVal action As Integer)

    PreConfiguration action

    Select Case action
        Case 0 : InitializationPhase     ' 扫参开始前: 清理环境、准备结果树
        Case 1 : ProcessingPhase         ' 每个相位点: 读场、积分、缓存结果行
        Case 2 : FinalizationPhase       ' 扫参结束: 汇总曲线、写入结果树
    End Select
End Sub

' 每次调用都先执行: 读配置、缓存工程环境、解析使能位与周期方向、确定参考位置
Private Sub PreConfiguration(ByVal action As Integer)
    Dim sMeshType As String

    LoadMacroConfig                          ' 版本信息必须来自 LinXi.ini, 失败即中止

    sMeshType = Mesh.GetMeshType

    ' 缓存本工程的环境全局变量
    g_bIsTetra    = (sMeshType = "Tetrahedral")
    g_sTemp       = GetProjectPath("Temp")
    g_sResult     = GetProjectPath("Result")
    g_dUnitToSI   = Units.GetGeometryUnitToSI
    g_sUnit       = Units.GetGeometryUnit
    g_sLogFile    = g_sTemp & LOG_NAME
    g_sPointsFile = g_sTemp & "_field_points.txt"
    g_iNumModes   = Solver.AKSGetNumberOfModes

    ' 文件名标签预算 = 总允许长度 - Temp 路径长度 - 固定开销, 防止超出 260 字符
    g_nFileTagBudget = MAX_PATH_SAFE - Len(g_sTemp) - FILE_TAG_FIXED

    If action = 0 Then
        InitLogFile
        ' 路径预算不足时必须在写盘前中止: 否则标签会超出总长, 临时文件写不进又清理不掉
        If g_nFileTagBudget < MIN_FILE_TAG_BUDGET Then
            LogCritical "工程路径过长: Temp 路径长度 = " & Len(g_sTemp) & _
                " 字符, 文件名标签仅剩 " & g_nFileTagBudget & " 字符 (至少需要 " & _
                MIN_FILE_TAG_BUDGET & " 字符)。请将工程移至更短的目录后重新运行", True
            End
        End If
    End If

    If Not g_bCfgOK Then AbortOnConfigError

    Dim Macro_flag As Double
    Dim bNeedSetupDialog As Boolean
    Dim iFlagSrc As Integer
    Dim iFlagRegion As Integer

    bNeedSetupDialog = False
    g_iRegionMode = REGION_BOTH

    ' 使能位 Macro_SweepWatch_Enable = 整数位(功率流方式) + 一位小数位(计算区域)
    If DoesParameterExist("Macro_SweepWatch_Enable") Then
        Macro_flag = RestoreParameter("Macro_SweepWatch_Enable")

        If Abs(Macro_flag) < 1E-09 Then
            SetParameterDescription("Macro_SweepWatch_Enable", "宏禁用状态")
            End
        ElseIf DecodeEnableFlag(Macro_flag, iFlagSrc, iFlagRegion) Then
            g_iPowerFlowSrc = iFlagSrc
            g_iRegionMode = iFlagRegion
            SetParameterDescription "Macro_SweepWatch_Enable", EnableFlagText(iFlagSrc, iFlagRegion)
        Else
            bNeedSetupDialog = True
            g_iPowerFlowSrc = PW_SRC_INDIRECT
            g_iRegionMode = REGION_BOTH
            LogWarning "  Macro_SweepWatch_Enable = " & Format(Macro_flag, "0.0####") & _
                " 不是合法的使能位取值, 本次运行将重新弹出参数设置对话框!"
        End If
    Else
        bNeedSetupDialog = True
        g_iPowerFlowSrc = PW_SRC_INDIRECT
        g_iRegionMode = REGION_BOTH
    End If

    ' 唯一的周期方向即扫描方向; 只允许存在 1 个周期方向, 不允许存在 0 个或 2 个及以上的周期方向
    With Boundary
        Dim nPeriodic As Integer
        nPeriodic = 0
        If .GetXmin = "periodic" Then g_iDir = 1 : nPeriodic = nPeriodic + 1
        If .GetYmin = "periodic" Then g_iDir = 2 : nPeriodic = nPeriodic + 1
        If .GetZmin = "periodic" Then g_iDir = 3 : nPeriodic = nPeriodic + 1
        If nPeriodic > 1 Then
            LogCritical("检测到多个周期方向!", True)
            End
        ElseIf g_iDir = 0 Then
            LogCritical("未设置周期性边界条件!", True)
            End
        End If
        ' 识别计算域空间; 周期方向长度即结构周期 pitch
        .GetCalculationBox g_cMin(1), g_cMax(1), _
            g_cMin(2), g_cMax(2), _
            g_cMin(3), g_cMax(3)
        g_dPitchUnit = g_cMax(g_iDir) - g_cMin(g_iDir)
        If g_dPitchUnit <= 0 Then
            LogCritical("周期长度为零, 可能网格数不足", True)
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

    ' 三个参考位置参数与至少一个 Macro_Modex 同时存在 => 视为已完成设置, 不再弹窗
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
        ' --- 分支一: 沿用既有设置, 由 Macro_Modex 反推参与计算的模式 ---
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
                    "Mode " & CStr(iChk) & " 耦合阻抗计算标志"
            End If
        Next iChk

        ' 全选与未选都归结为"全部模式"
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
        ' --- 分支二: 首次运行或设置被删除, 弹出参数设置对话框 ---
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
            ' 用户取消: 关掉使能位, 本次宏不运行, 但 CST 扫参继续
            StoreParameter("Macro_SweepWatch_Enable", 0#)
            SetParameterDescription("Macro_SweepWatch_Enable", "宏禁用状态")
            LogInfo "用户取消了宏程序参数设置! 使能标志位变更为0, 宏程序将不会运行! 但CST扫参程序仍会进行!"
            End
        End If

        ' 取值 = 功率流方式(1/2) + 计算区域(0/1/2) / 10
        StoreParameter "Macro_SweepWatch_Enable", _
            CDbl(g_iPowerFlowSrc + 1) + CDbl(g_iRegionMode) / 10#
        SetParameterDescription "Macro_SweepWatch_Enable", _
            EnableFlagText(g_iPowerFlowSrc, g_iRegionMode)

        ' 周期方向的参考位置强制取计算域中心
        pos(g_iDir) = 0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir))
        Dim iPos As Integer
        For iPos = 1 To 3
            If Abs(pos(iPos)) < 1e-12 Then pos(iPos) = 0#
        Next iPos

        StoreParameter "Kc_RefPos_x", pos(1)
        StoreParameter "Kc_RefPos_y", pos(2)
        StoreParameter "Kc_RefPos_z", pos(3)
        SetParameterDescription "Kc_RefPos_x", "耦合阻抗参考位置的X坐标"
        SetParameterDescription "Kc_RefPos_y", "耦合阻抗参考位置的Y坐标"
        SetParameterDescription "Kc_RefPos_z", "耦合阻抗参考位置的Z坐标"

        ' 先清理残留的 Macro_Modex, 再按本次选择重建 (参数"存在"即代表选定对应的模式)
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
                    "Mode " & CStr(iChk) & " 耦合阻抗计算标志"
            Next iChk
        Else
            For iIdx = 0 To g_nSelectedModes - 1
                iChk = g_aSelectedModes(iIdx)
                StoreParameter "Macro_Mode" & CStr(iChk), 0#
                SetParameterDescription "Macro_Mode" & CStr(iChk), _
                    "Mode " & CStr(iChk) & " 耦合阻抗计算标志"
            Next iIdx
        End If
    End If

    g_xRef = pos(1)
    g_yRef = pos(2)
    g_zRef = pos(3)

    ' 非周期方向的参考位置必须落在与周期方向垂直的截面内部, 否则无法积分或积分错误
    Select Case g_sDirLabel
        Case "X"
            If g_yRef <= g_cmin(2) Or g_yRef >= g_cmax(2) Then LogCritical("耦合阻抗参考位置的 Y 坐标落在计算域外!", True)
            If g_zRef <= g_cmin(3) Or g_zRef >= g_cmax(3) Then LogCritical("耦合阻抗参考位置的 Z 坐标落在计算域外!", True)
        Case "Y"
            If g_xRef <= g_cmin(1) Or g_xRef >= g_cmax(1) Then LogCritical("耦合阻抗参考位置的 X 坐标落在计算域外!", True)
            If g_zRef <= g_cmin(3) Or g_zRef >= g_cmax(3) Then LogCritical("耦合阻抗参考位置的 Z 坐标落在计算域外!", True)
        Case "Z"
            If g_xRef <= g_cmin(1) Or g_xRef >= g_cmax(1) Then LogCritical("耦合阻抗参考位置的 X 坐标落在计算域外!", True)
            If g_yRef <= g_cmin(2) Or g_yRef >= g_cmax(2) Then LogCritical("耦合阻抗参考位置的 Y 坐标落在计算域外!", True)
    End Select

    BuildModeFlagTable
End Sub

' 扫参开始前的一次性准备: 清空上一轮的临时文件与结果树, 复位各状态数组
Private Sub InitializationPhase()

    g_sLastGroupKey = ""
    On Error Resume Next
    Kill g_sTemp & "_groups.txt"
    On Error GoTo 0

    ' 删除上一次运行留下的结果分组 (只删该分组, 不动其它结果)
    DeleteTreeItemRecursive "1D Results\" & F_RESULT_GROUP

    ' 防御性代码, CST 本征模求解器模式数量至少为 1
    If g_iNumModes = 0 Then
        LogCritical "未找到任何本征模式", True
        End
    End If

    ' 清理 Temp 目录内上一次宏产生的全部中间文件
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

    ' 清理超出当前模式数的 Macro_Mode 残留参数
    Dim iCleanup As Integer
    For iCleanup = g_iNumModes + 1 To MAX_MODE_PARAMS
        On Error Resume Next
        If DoesParameterExist("Macro_Mode" & CStr(iCleanup)) Then
            DeleteParameter "Macro_Mode" & CStr(iCleanup)
        End If
        On Error GoTo 0
    Next iCleanup

    ' 周期方向的参考坐标 = 计算域中心
    Select Case g_sDirLabel
        Case "X"
            StoreParameter "Kc_RefPos_x", IIf(Abs(0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir))) < 1e-12, 0, 0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)))
            SetParameterDescription "Kc_RefPos_x", "耦合阻抗参考位置的X坐标, 周期方向"
            SetParameterDescription "Kc_RefPos_y", "耦合阻抗参考位置的Y坐标"
            SetParameterDescription "Kc_RefPos_z", "耦合阻抗参考位置的Z坐标"
        Case "Y"
            StoreParameter "Kc_RefPos_y", IIf(Abs(0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir))) < 1e-12, 0, 0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)))
            SetParameterDescription "Kc_RefPos_x", "耦合阻抗参考位置的X坐标"
            SetParameterDescription "Kc_RefPos_y", "耦合阻抗参考位置的Y坐标, 周期方向"
            SetParameterDescription "Kc_RefPos_z", "耦合阻抗参考位置的Z坐标"
        Case "Z"
            StoreParameter "Kc_RefPos_z", IIf(Abs(0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir))) < 1e-12, 0, 0.5 * (g_cMin(g_iDir) + g_cMax(g_iDir)))
            SetParameterDescription "Kc_RefPos_x", "耦合阻抗参考位置的X坐标"
            SetParameterDescription "Kc_RefPos_y", "耦合阻抗参考位置的Y坐标"
            SetParameterDescription "Kc_RefPos_z", "耦合阻抗参考位置的Z坐标, 周期方向"
    End Select
End Sub

' 单个扫描点的主流程: 取相位 -> 读频率 -> 判有效性 -> 读场 -> 积分 -> 缓存结果行
Private Sub ProcessingPhase()

    Dim dPhaseDeg     As Double
    Dim dPhaseRad     As Double
    Dim dBeta         As Double

    ' 按需扩容结果行缓冲与相位状态数组
    If g_iNumModes > 0 And g_nBufModes <> g_iNumModes Then
        ReDim g_aBufBeta(1 To g_iNumModes)
        ReDim g_aBufZpierce(1 To g_iNumModes)
        ReDim g_aBufPhase(1 To g_iNumModes)
        g_nBufModes = g_iNumModes
    End If

    If g_iNumModes > 0 And g_nPhaseStateModes <> g_iNumModes Then
        ReDim g_abPhaseDead(1 To g_iNumModes)
        ReDim g_anPhaseFlat(1 To g_iNumModes)
        ReDim g_abPhaseFlatWarned(1 To g_iNumModes)
        ReDim g_abKcWarned(1 To g_iNumModes)
        g_nPhaseStateModes = g_iNumModes
    End If

    Dim nParams     As Long
    Dim i           As Integer
    Dim nPhaseIndex As Integer
    Dim bFound      As Boolean
    Dim sParamName  As String
    Dim dParamVal   As Double

    ' 扫描列表中必须存在 phase; 其余变量共同构成 "参数组合" 分组键
    With ParameterSweep
        nParams = .GetNumberOfVaryingParameters

        bFound       = False
        nPhaseIndex  = -1
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
            LogCritical "参数列表中未找到 phase 参数!", True
            End
        End If

        dPhaseDeg = .GetValueOfVaryingParameter(nPhaseIndex)
    End With

    ' 每个参数组合分配一个稳定序号, 保证结果树分组与文件名一一对应
    Dim nGroupIndex As Long, nGroupCount As Long
    nGroupIndex = LookupGroupIndex(g_sGroupKey, nGroupCount)
    If nGroupIndex = 0 Then nGroupIndex = nGroupCount + 1

    Dim fGroup As Integer
    fGroup = FreeFile
    Open g_sTemp & "_groups.txt" For Append As #fGroup
    Print #fGroup, g_sGroupPath & "|" & g_sGroupKey & "|" & nGroupIndex
    Close #fGroup

    ' 文件名标签: 组合键过长或含非法字符时退化为"截断 + _g序号"
    g_sFileTag = MakeFileTag(g_sGroupKey, nGroupIndex, g_nFileTagBudget)

    Dim sFileSuffix As String
    If g_sFileTag = "" Then sFileSuffix = "" Else sFileSuffix = "_" & g_sFileTag

    ' 相位差 -> 相位常数 beta = phi / pitch
    dPhaseRad = dPhaseDeg * DegToRad
    dBeta = dPhaseRad / g_dPitchSI

    If g_sGroupKey <> "" Then LogInfo " ---------- 参数组合: " & g_sGroupKey & " ---------- "

    LogInfo " ----- phase = " & dPhaseDeg & "° ----- "

    ' 读取上一扫描点各模式的频率, 用于判断频率是否随相位递增
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

    ' phase 为 Pi 的整数倍时, cos(phi) 位置的特殊性使耦抗计算无意义, 直接跳过
    Dim bPhaseIsPiMultiple As Boolean
    bPhaseIsPiMultiple = (Abs(Sin(dPhaseDeg * DegToRad)) < 1e-4)

    Dim sMode As String
    Dim dFreqHz As Double, dFreqGHz As Double
    Dim dVre As Double, dVim As Double, dEabs As Double, dVSq As Double
    Dim dPower As Double, dKc As Double

    ' 进入新的参数组合: 复位采样网格与相位状态
    If g_sGroupKey <> g_sLastGroupKey Then
        g_nCross = 0
        g_bGridFailed = False
        If RowBuffersPending Then
            LogError "参数组合切换: 丢弃上一组合未能写出的临时数据行"
        End If
        ClearRowBuffers
        ResetPhaseSweepState
        g_sLastGroupKey = g_sGroupKey
    End If

    ' 采样网格只与几何有关, 每个参数组合计算一次后复用
    If Not bPhaseIsPiMultiple Then
        PrecomputeMeshGrid
    Else
        LogInfo "  ***  相位 " & dPhaseDeg & "° 为 Pi 整数倍  ***"
        LogInfo "  ***  跳过 采样网格 / 点文件 的生成, 跳过 场积分 / 功率流 / 耦合阻抗 的计算  ***"
    End If

    Mesh.ViewMeshMode False

    For iMode = 1 To g_iNumModes
        dFreqHz = 0#

        sMode = CStr(iMode)

        Dim bModeSelected As Boolean
        bModeSelected = g_abModeFlag(iMode)
        If Not bModeSelected Then
            LogInfo "  未选定 Mode " & sMode & " , 跳过"
            GoTo SkipModeCalc
        End If

        ' 取本征模频率: 先走 Result3D 直读, 失败再选中结果树后重试
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
            LogError "  相位 " & dPhaseDeg & "° 下 Mode " & sMode & _
                " 读取场频率失败 (已重试 " & iRetry & " 次), 本点跳过"
            GoTo SkipModeCalc
        End If

        dFreqGHz = dFreqHz * Units.GetFrequencyUnitToSI / 1e9
        dCurFreq(iMode) = dFreqHz
        abFreqOK(iMode) = True

        ' 频率升降判据: 行波区要求频率随 phase 上升, 返波区要求下降, 双区模式不作判断
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

        ' 频率有效性判定, 连续 FREQ_FLAT_MAX 点即判整个相位扫描无效
        If bFreqFlat Then
            g_anPhaseFlat(iMode) = g_anPhaseFlat(iMode) + 1
            If Not g_abPhaseFlatWarned(iMode) Then
                g_abPhaseFlatWarned(iMode) = True
                LogWarning "  Mode " & sMode & ": 频率未随 phase 变化 (f = " & _
                    Format(dFreqHz, "0.000E+00") & " " & Units.GetFrequencyUnit & ", 相对变化 " & _
                    Format(dFreqRel, "0.0E+00") & ")。 原因通常是: 边界设置中没有把 phase 施加到 " & _
                    g_sDirLabel & " 方向周期边界的相位差上, 或该模式位于 Pi 点附近。 两种情况下的净功率流都接近零, 耦合阻抗无意义, 已判定为无效!"
            ElseIf g_anPhaseFlat(iMode) <= FREQ_FLAT_MAX Then
                LogInfo "  Mode " & sMode & ": 频率仍未随 phase 变化 (连续 " & g_anPhaseFlat(iMode) & " 点), 本扫描点耦合阻抗判定为无效!"
            End If
            If g_anPhaseFlat(iMode) >= FREQ_FLAT_MAX And Not g_abPhaseDead(iMode) Then
                g_abPhaseDead(iMode) = True
                LogWarning "  Mode " & sMode & ": 连续 " & FREQ_FLAT_MAX & " 个扫描点的频率完全相同, " & _
                    "判定本次扫描的相位扫描无效, 请检查 " & g_sDirLabel & " 方向: " & _
                    "(1) 两侧边界是否都设为 periodic; (2) 周期边界的相位差是否绑定到扫描变量 phase; " & _
                    "(3) 计算域在该方向是否恰好为一个周期。"
                LogWarning "  Mode " & sMode & ": 后续扫描点将跳过场导出与积分, 仅记录场频率与相位"
            End If
        Else
            g_anPhaseFlat(iMode) = 0
            If g_abPhaseDead(iMode) Then
                g_abPhaseDead(iMode) = False
                g_abPhaseFlatWarned(iMode) = False
                LogInfo "  Mode " & sMode & ": 频率重新随 phase 变化, 恢复该模式的耦合阻抗计算"
            End If
        End If

        ' 三种情况跳过后处理: Pi 整数倍点 / 频率变化方向与所选区域不符 / 相位扫描已判无效
        Dim bSkipImpedance As Boolean
        bSkipImpedance = bPhaseIsPiMultiple Or (Not bFreqDirectionOK) Or g_abPhaseDead(iMode)

        If bSkipImpedance Then
            ' 无效标记统一用 -1, 汇总阶段据此剔除该点
            dVre = 0#: dVim = 0#: dEabs = 0#: dVSq = 0#
            dPower = -1#
            dKc = -1#

            LogInfo "  Mode " & sMode & ": freq = " & Format(dFreqGHz, "0.000000") & " GHz"

            If Not bPhaseIsPiMultiple And Not bFreqFlat Then
                If g_abPhaseDead(iMode) Then
                    LogInfo "  Mode " & sMode & ": 已判定相位扫描无效, 跳过 场导出 / 场积分 / 功率流 / 耦合阻抗"
                ElseIf Not bFreqDirectionOK Then
                    Select Case g_iRegionMode
                        Case REGION_FORWARD
                            LogInfo "  *** 频率未上升: " & Format(dPrevFreq(iMode), "0.000E+00") & _
                                " → " & Format(dFreqHz, "0.000E+00") & " " & Units.GetFrequencyUnit & _
                                ", 不属于行波区, 跳过 场积分 / 功率流 / 耦合阻抗 的计算  ***"
                        Case REGION_BACKWARD
                            LogInfo "  *** 频率未下降: " & Format(dPrevFreq(iMode), "0.000E+00") & _
                                " → " & Format(dFreqHz, "0.000E+00") & " " & Units.GetFrequencyUnit & _
                                ", 不属于返波区, 跳过 场积分 / 功率流 / 耦合阻抗 的计算  ***"
                    End Select
                End If
            End If
        Else
            ' ---- 正常路径: 导出采样点场数据 ----
            LogInfo "  Mode " & sMode & ": freq = " & Format(dFreqGHz, "0.000000") & _
                " GHz, 开始计算..."

            Dim bFieldsOK As Boolean
            If g_iPowerFlowSrc = PW_SRC_NATIVE Then
                bFieldsOK = PreparePowerFieldFromCST(iMode, sMode)   ' 原生法: 由 CST 叉积生成功率流场
            Else
                bFieldsOK = LoadFieldDataFromEH(iMode)               ' 间接法: 导出 E 与 H 场
            End If

            If bFieldsOK Then
                ' 沿周期方向积分得到等效电压 V (复数量), |V|^2 用于耦抗公式
                IntegrateLongitudinal dBeta, dVre, dVim, dEabs
                dVSq = dVre * dVre + dVim * dVim
                LogInfo "    纵向积分: V_avg = (" & Format(dVre, "0.000E+00") & ", " & _
                    Format(dVim, "0.000E+00") & "), |V|^2 = " & Format(dVSq, "0.000E+00") & _
                    ", E_abs = " & Format(dEabs, "0.000E+00")

                ' 沿横截面积分得到净功率流 P
                If g_iPowerFlowSrc = PW_SRC_NATIVE Then
                    ComputePowerFlowFromCST dPower
                Else
                    ComputePowerFlowFromEH dPower
                End If
                LogInfo "    功率流: P = " & Format(dPower, "0.000E+00") & " W"

                If g_iPowerFlowSrc = PW_SRC_INDIRECT And dPower = 0# Then
                    LogWarning "    提示: E/H 方式下功率流为 0。请确认结果树中 Mode " & sMode & _
                        " 的磁场项可正常选取!"
                End If

                ' 皮尔斯耦合阻抗 Kc = |V|^2 / (2 * beta^2 * P); 分母异常或结果越界时记为无效
                If dVSq = 0# Then
                    dKc = -1#
                    LogWarning "    纵向积分无有效结果 (|V|^2 = 0), 本点耦合阻抗记为无效!"
                ElseIf dBeta = 0# Or dPower <= 0# Then
                    dKc = -1#
                    If dBeta = 0# Then
                        LogWarning "    相位常数为零 (beta = 0), 本点耦合阻抗记为无效!"
                    Else
                        LogWarning "    功率流非正 (P = " & Format(dPower, "0.000E+00") & " W), 本点耦合阻抗记为无效!"
                    End If
                Else
                    dKc = dVSq / (2# * dBeta * dBeta * dPower)
                    If dKc > KC_SANITY_MAX Then
                        If Not g_abKcWarned(iMode) Then
                            g_abKcWarned(iMode) = True
                            LogWarning "    功率流异常小: P = " & Format(dPower, "0.000E+00") & " W, 对应 Kc = " & _
                                Format(dKc, "0.000E+00") & " Ohm, 超过合理上限 " & _
                                Format(KC_SANITY_MAX, "0.000E+00") & " Ohm, 本点耦合阻抗记为无效!"
                        Else
                            LogInfo "    Kc 超过合理上限 (" & Format(dKc, "0.000E+00") & " Ohm), 本点记为无效!"
                        End If
                        dKc = -1#
                    End If
                End If
                LogInfo "    耦合阻抗: Kc = " & Format(dKc, "0.000E+00") & " Ohm"
            Else
                LogError "Mode " & sMode & ": 场导出 / 解析失败, 本点输出无效值!"
                dVre = 0#: dVim = 0#: dEabs = 0#: dVSq = 0#
                dPower = -1#: dKc = -1#
            End If
        End If

        ' 本轮扫描点结果先写入内存缓冲 (频率, 值)
        BufferAppend g_aBufBeta(iMode),    dFreqGHz, dBeta
        BufferAppend g_aBufZpierce(iMode), dFreqGHz, dKc
        BufferAppend g_aBufPhase(iMode),   dFreqGHz, dPhaseDeg
SkipModeCalc:
    Next iMode

    ' 各模式的缓冲一次性写出到临时 sig 文件
    For iMode = 1 To g_iNumModes
        bModeSelected = g_abModeFlag(iMode)
        If bModeSelected Then
            BufferFlush TempPath(F_BETA    & sFileSuffix, iMode), g_aBufBeta(iMode)
            BufferFlush TempPath(F_ZPIERCE & sFileSuffix, iMode), g_aBufZpierce(iMode)
            BufferFlush TempPath(F_PHASE   & sFileSuffix, iMode), g_aBufPhase(iMode)
        End If
    Next iMode

    ' 记录本点频率, 供下一个扫描点做递增/冻结判断
    For iMode = 1 To g_iNumModes
        If g_abModeFlag(iMode) And abFreqOK(iMode) Then
            iPrevFile = FreeFile
            Open g_sTemp & "_prevfreq_" & iMode & sFileSuffix & ".txt" For Output As #iPrevFile
            Print #iPrevFile, dCurFreq(iMode)
            Close #iPrevFile
        End If
    Next iMode

End Sub

' 扫描结束后的汇总: 逐参数组合、逐模式生成曲线并挂到结果树, 然后清理临时文件
Private Sub FinalizationPhase()
    LogInfo "======================================================================="
    LogInfo "  扫描结束, 开始汇总结果..."
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

    ' 回读 _groups.txt (每行: 结果树路径|组合键|组合序号), 按路径去重得到参数组合清单
    On Error Resume Next
    If FileExists(g_sTemp & "_groups.txt") Then
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

                ' 同一参数组合在扫描中被写入多次, 这里只保留首次出现
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

    ' 第一轮: 为每个组合的每个模式生成四类曲线并挂到结果树
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
                LogInfo "  Mode " & sMode & ": 未选定, 跳过处理"
                GoTo SkipModeFinalize
            End If

            If sPath = "" Then
                LogInfo "---------- Mode " & sMode & ": 汇总结果曲线 ----------"
            Else
                LogInfo "---------- 参数组合 [ " & Replace(sPath, "\", " & ") & " ] Mode " & sMode & ": 汇总结果曲线 ----------"
            End If

            Dim sCurBeta As String, sCurPhase As String, sCurVp As String, sCurKc As String
            sCurBeta  = g_sTemp & "_curveBeta"  & sFileSuffix2 & "_" & sMode & ".txt"
            sCurPhase = g_sTemp & "_curvePhase" & sFileSuffix2 & "_" & sMode & ".txt"
            sCurVp    = g_sTemp & "_curveVp"    & sFileSuffix2 & "_" & sMode & ".txt"
            sCurKc    = g_sTemp & "_curveKc"    & sFileSuffix2 & "_" & sMode & ".txt"

            ' 由该模式本组合的临时数据生成四条曲线文件, 并取回各曲线有效点数
            Dim nBetaPt As Long, nPhasePt As Long, nVpPt As Long, nKcPt As Long
            Dim nRetCurve As Long
            nRetCurve = CBuildCurveFiles(CORE_AUTHOR, g_sTemp, sFileSuffix2, iMode, CLight, _
                sCurBeta, sCurPhase, sCurVp, sCurKc, _
                nBetaPt, nPhasePt, nVpPt, nKcPt)

            If nRetCurve <> 0 Then
                LogError "Mode " & sMode & ": 无有效扫描数据 (曲线数据生成失败, 返回码 " & nRetCurve & "), 跳过本模式"
                GoTo SkipModeFinalize
            End If

            ' 耦合阻抗在部分扫描点被记为无效, 点数少于 beta 属正常情况
            If nKcPt < nBetaPt Then
                LogInfo "  Mode " & sMode & ": " & (nBetaPt - nKcPt) & _
                    " 个扫描点的耦合阻抗无效 (可能是 |V|^2 无结果, 或频率变化方向与所选计算区域不符), 已跳过"
            End If

            ' 以下四条曲线按"有数据才保存"的原则各自独立写入结果树
            If nBetaPt > 0 Then
                Set oBeta = Result1D("")
                oBeta.LoadPlainFile sCurBeta
                Save1D oBeta, "Brillouin Diagram", _
                    g_sResult & F_BETA & "_" & sMode & sFileSuffix2 & ".sig", _
                    sTreeBase & "\Brillouin Diagram Beta\Mode " & sMode, _
                    "Beta (rad/m)", "Frequency (GHz)"
                LogInfo "  Mode " & sMode & ": 布里渊曲线 Beta 已保存"
            End If

            If nPhasePt > 0 Then
                Set oBetaPhase = Result1D("")
                oBetaPhase.LoadPlainFile sCurPhase
                Save1D oBetaPhase, "Brillouin Diagram (Phase)", _
                    g_sResult & F_PHASE & "_" & sMode & sFileSuffix2 & ".sig", _
                    sTreeBase & "\Brillouin Diagram Phase\Mode " & sMode, _
                    "Phase (deg)", "Frequency (GHz)"
                LogInfo "  Mode " & sMode & ": 布里渊曲线 Phase 已保存"
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

    ' 第二轮: 结果已全部入库, 清理各类临时文件
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
        LogInfo "  Mode " & iMode & ": 临时文件已清理"
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
        If nDay > 0 Then sTime = sTime & nDay & "天"
        If nDay > 0 Or nHour > 0 Then sTime = sTime & nHour & "时"
        If nDay > 0 Or nHour > 0 Or nMin > 0 Then sTime = sTime & nMin & "分"
        sTime = sTime & nSec & "秒"
    Else
        sTime = "(缺失开始时刻记录, 无法计算运行时间)"
    End If

    LogInfo "======================================================================="
    LogInfo "  全部完成! 共处理 " & nG & " 个参数组合 × " & g_iNumModes & " 个模式"
    LogInfo "  耦合阻抗参考位置: X = " & g_xRef & ", Y = " & g_yRef & _
        ", Z = " & g_zRef & ",  单位: " & g_sUnit & ", 周期方向: " & g_sDirLabel
    LogInfo "  功率流计算方式: " & PowerFlowSrcText(g_iPowerFlowSrc) & _
        ",  计算区域: " & RegionText(g_iRegionMode)
    LogInfo "  日志文件: " & g_sLogFile
    If Len(g_sLogLastError) > 0 Then
        LogWarning "  本次运行出现过日志写盘失败: " & g_sLogLastError
    End If
    LogInfo "  运行时间: " & sTime
    LogInfo "  结束时间: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LogInfo "======================================================================="

End Sub

' 读取 LinXi.ini 并校验 [Version] 节; 任一环节失败都只记录原因, 由中止逻辑统一提示
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
        g_sCfgNote = "4 个候选位置均未找到外部配置文件 " & CFG_FILE_NAME
    ElseIf Not ReadConfigFile(g_sCfgPath) Then
        g_sCfgNote = "外部配置文件无法读取(可能被占用或权限不足): " & g_sCfgPath
    Else
        sMiss = CfgMissingKeys()
        If Len(sMiss) > 0 Then
            g_sCfgNote = "配置文件缺少 [Version] 节必填键: " & sMiss & " —— " & g_sCfgPath
        ElseIf Not ParseDllMinVersion() Then
            g_sCfgNote = "配置文件中的 DllMinVersion 不是有效的版本号: """ & _
                g_sCfgDllMinRaw & """ (应为正整数, 例如 122) —— " & g_sCfgPath
        Else
            g_sVersionString = g_sCfgVerRaw
            g_sVersionLabel = g_sCfgLblRaw
            g_bCfgOK = True
        End If
    End If
End Sub

' 校验 DllMinVersion 为正整数并写入 g_nDllMinVer
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

' 返回 [Version] 节中缺失的必填键列表 (空串表示齐全)
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

' 配置文件的第 iIndex 个候选路径 (安装目录/宏目录/工程目录), 取不到时返回空串
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

' 文件是否真实存在 (Dir 失败一律按不存在处理)
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

' 按候选顺序定位 LinXi.ini; 调用方据此给出"4 个位置都没找到"的提示
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

' 逐行解析 ini: 支持 ; 与 # 注释、[节] 与 key=value, 命中键交给 ApplyConfigKey
' 注意: 必须为 ANSI/GBK 编码, UTF-8 存档会导致所有键读不到
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

    ' 打开失败(被占用/权限不足)直接失败返回
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
                ' 整行注释, 跳过
            ElseIf Left$(sLine, 1) = "[" Then
                ' 节名统一转大写比较
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

' 把 [Version] 节的一个键值写入对应的全局变量 (节名或键名不匹配则忽略)
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

' 配置不可用时的终止处理: 弹窗列出已查找的 4 个候选位置后中止宏
Private Sub AbortOnConfigError()
    Dim sMsg As String
    Dim i As Integer

    sMsg = "错误: 宏程序的版本信息必须来自外部配置文件 " & CFG_FILE_NAME & _
        ", 但当前无法取得, 宏程序已中止!" & vbCrLf & vbCrLf & _
        "原因: " & g_sCfgNote & vbCrLf & vbCrLf & _
        "宏按下列顺序查找该文件, 请确认至少有一处存在且可读, 且 [Version] 节的" & vbCrLf & _
        "8 个键齐全(代码内已不保留兜底默认值):" & vbCrLf
    For i = 1 To 4
        If Len(CfgCandidate(i)) > 0 Then sMsg = sMsg & "  " & i & ". " & CfgCandidate(i) & vbCrLf
    Next i
    sMsg = sMsg & vbCrLf & _
        "提示: 该文件必须保存为 ANSI/GBK 编码; 若被另存为 UTF-8, 将读不到任何键。" & vbCrLf & _
        "      重新运行安装脚本(双击我自动安装.bat)可恢复默认部署。"

    MsgBox sMsg, vbCritical, "宏程序已中止"
    On Error Resume Next
    LogCritical sMsg, True
    On Error GoTo 0
    End
End Sub

' 校验核心计算库: 文件是否存在、CCoreVersion 能否调用、版本是否满足最低要求
Private Sub CheckCoreLibrary()
    Dim sDllPath As String
    sDllPath = GetInstallPath & "\AMD64\LinXi.dll"

    If Not FileExists(sDllPath) Then
        MsgBox "错误: 无法加载核心计算库 LinXi.dll。" & vbCrLf & _
            "未找到核心计算库文件:" & vbCrLf & sDllPath & vbCrLf & vbCrLf & _
            "请确认:" & vbCrLf & _
            "  1. 已正确部署核心计算库 LinXi.dll;" & vbCrLf & _
            "  2. CST 程序为 64 位版本 (DLL 为 64 位版本)。", vbCritical
        LogCritical "核心计算库 LinXi.dll 加载失败, 文件不存在: " & sDllPath, True
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
        MsgBox "错误: 核心计算库加载失败。" & vbCrLf & vbCrLf & _
            "文件存在, 但无法正常调用。可能原因:" & vbCrLf & _
            "  1. DLL 是 32 位（应为 64 位）;" & vbCrLf & _
            "  2. 导出函数名不匹配;" & vbCrLf & _
            "  3. DLL 依赖的其他库缺失, 请检查或更新 DLL 版本。" & vbCrLf & vbCrLf & _
            "当前返回版本: v" & nVer & ", 本程序要求不低于 v" & g_nDllMinVer & vbCrLf & _
            "(最低版本要求写在 " & CFG_FILE_NAME & " 的 DllMinVersion)" & vbCrLf & vbCrLf & _
            "Err.Number = " & nErrNum & vbCrLf & _
            "LastDLLError = " & nDllErr & vbCrLf & _
            "Description = " & sDesc, vbCritical
        LogCritical " -----  核心计算库加载失败: Err=" & nErrNum & ", LastDLLError=" & nDllErr & _
            ", 返回版本=" & nVer & ", 要求不低于=" & g_nDllMinVer & "  ----- ", True
        End
    End If

    LogInfo " ----- 核心计算库 LinXi.dll 加载成功 v" & nVer & " ----- "

    If g_bCfgOK Then
        LogInfo " ----- CST 本征模慢波结构用户监视器宏代码程序 " & g_sVersionString & " (" & g_sVersionLabel & ")" &" ----- "
    Else
        LogWarning " ----- 版本信息未能从 " & CFG_FILE_NAME & " 读入: " & g_sCfgNote & " ----- "
    End If

End Sub

' 参数设置对话框的事件回调
' Action=1 对话框初始化; Action=2 控件事件 (DlgItem 为控件名, SuppValue 为控件值)
Private Function DialogFunc(ByVal DlgItem$, ByVal Action%, ByVal SuppValue&) As Boolean
    Select Case Action
        Case 1
            ' --- 初始化: 先全部禁用, 再按当前选项放开 ---
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

            ' 周期方向的参考坐标锁定为计算域中心, 界面上置灰显示
            Select Case g_sDirLabel
                Case "X"
                    DlgEnable "KcPosx", 0
                    DlgText   "KcPosx", "计算域中心"
                Case "Y"
                    DlgEnable "KcPosy", 0
                    DlgText   "KcPosy", "计算域中心"
                Case "Z"
                    DlgEnable "KcPosz", 0
                    DlgText   "KcPosz", "计算域中心"
            End Select

            ' 按模式选择方式 (全部 / 勾选 / 自定义) 展示对应控件
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
                    DlgText   "lblHint", "温馨提示: 请勾选需要计算耦合阻抗的模式"
                Case 2
                    DlgVisible "txtMode",     1
                    DlgVisible "lblHint",     1
                    DlgText   "lblHint", "温馨提示: 用英文逗号分隔, 例如 1,3,5,7"
                    DlgEnable "txtMode", 1
                    DlgFocus  "txtMode"
            End Select

            ' 功率流方式改变时切换对应的说明文字
            If DlgValue("Group2") = 1 Then
                DlgVisible "lblPwDesc1", 0
                DlgVisible "lblPwDesc2", 1
            Else
                DlgVisible "lblPwDesc1", 1
                DlgVisible "lblPwDesc2", 0
            End If

            ' 只显示当前所选计算区域对应的说明文字
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
            ' --- 控件事件: 分支互斥, 每个分支结束后回传是否已处理 ---
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
                            DlgText   "lblHint", "温馨提示: 请勾选需要计算耦合阻抗的模式"
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
                            DlgText   "lblHint", "温馨提示: 用英文逗号分隔, 例如 1,3,5,7"
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
                    If DlgValue("Group1") <> 1 Then DlgValue "Group1", 1
                    DialogFunc = True
                Case "Group2"
                    ' 切换功率流方式时同步切换说明文字
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
                    ' 只显示当前所选区域对应的说明文字
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
                    ' 非模态子对话框: 不改变主对话框的返回值
                    ShowAboutDialog
                    DialogFunc = True
                Case "btnCancel"
                    ' 取消时明确告知: 只关闭宏, 不中断 CST 扫参
                    MsgBox "用户取消了宏程序参数设置!" & vbCrLf & vbCrLf & _
                        "使能标志位变更为0, 宏程序将不会运行!" & vbCrLf & vbCrLf & _
                        "但CST扫参程序仍会进行!", _
                        vbCritical, "参数设置取消"
                    DialogFunc = False
                Case "help"
                    ShowHelpDialog
                    DialogFunc = True
                Case Else
                    DialogFunc = False
            End Select
        Case 4
            ' Action=4: 对话框已关闭, 无需处理
    End Select
End Function

' 参数设置对话框: 确认参与计算的模式、耦抗参考位置、功率流方式与计算区域
' 返回值 True = 用户确认; False = 用户取消
Private Function ShowParamsDialog(ByRef pos() As Double, _
    ByRef bAllModes As Boolean) As Boolean

    Begin Dialog UserDialog 1080, 578, "CST 本征模慢波结构用户监视器宏代码程序 版本 " & g_sVersionString & " (" & g_sVersionLabel & ")" , .DialogFunc

        GroupBox 20, 7, 1040, 153, "[ 第一步 ] 确认需要计算耦合阻抗的模式编号   ---   已设置待求解的本征模数量: " & g_iNumModes & " 个"

        OptionGroup .Group1
            OptionButton 60, 28,  150, 14, " 计算全部模式", .allMode
            OptionButton 60, 50,  200, 14, " 仅计算选定的模式", .someMode
            OptionButton 60, 118, 130, 14, " 自定义模式", .someMode2

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

        GroupBox 20, 167, 1040, 69, "[ 第二步 ] 确认耦合阻抗参考位置计算点"

        Text  60, 188, 560, 14, "耦合阻抗计算点默认坐标为   X = 0  Y = 0  Z = 0   单位: " & g_sUnit

        Text     60, 212, 140, 14, "参考点 X 的坐标值"
        TextBox 210, 210, 140, 18, .KcPosx
        Text    400, 212, 140, 14, "参考点 Y 的坐标值"
        TextBox 550, 210, 140, 18, .KcPosy
        Text    740, 212, 140, 14, "参考点 Z 的坐标值"
        TextBox 890, 210, 140, 18, .KcPosz

        GroupBox 20, 243, 1040, 65, "[ 第三步 ] 确认功率流的计算方式"

        OptionGroup .Group2
            OptionButton 60, 264, 320, 14, "由电场 E 与磁场 H 计算得出功率流", .pwFromEH
            OptionButton 60, 286, 320, 14, "由 CST 原生生成功率流场", .pwFromCST

        Text 400, 264, 420, 14, "间接功率流法, 先插值, 后叉积, 再积分, 功率流为近似的乘积", .lblPwDesc1
        Text 400, 286, 420, 14, "原生功率流法, 先叉积, 后插值, 再积分, 功率流为乘积的近似", .lblPwDesc2

        GroupBox 20, 315, 1040, 87, "[ 第四步 ] 确认需要计算耦合阻抗的区域"

        OptionGroup .Group3
            OptionButton 60, 336, 320, 14, "计算行波区和返波区", .regionBoth
            OptionButton 60, 358, 320, 14, "仅计算行波区", .regionFwd
            OptionButton 60, 380, 320, 14, "仅计算返波区", .regionBwd

        Text 400, 336, 640, 14, "不判断频率随 phase 的升降, 所有扫描点都计算耦合阻抗", .lblRgnDesc0
        Text 400, 358, 640, 14, "行波区: 仅当频率随 phase 上升时计算, 频率下降的扫描点跳过", .lblRgnDesc1
        Text 400, 380, 640, 14, "返波区: 仅当频率随 phase 下降时计算, 频率上升的扫描点跳过", .lblRgnDesc2

        GroupBox 20, 409, 1040, 109, "[ 温馨提示 ]"

        Text 60, 430, 990, 14, "1. 宏程序仅检查参数列表中是否定义了 Macro_Modex, 以此决定对应模式的耦合阻抗是否参与计算；参数的具体取值不影响判断。"
        Text 60, 452, 990, 14, "2. 首次运行时会弹出此对话框, 由用户确认需计算的模式编号；后续宏程序通过检测参数列表中的 Macro_Modex 自动匹配计算。"
        Text 60, 474, 990, 14, "3. 宏程序通过 Macro_SweepWatch_Enable 的整数位识别功率流方式, 小数位识别计算区域: x.0 为两个区域, x.1 为行波区, x.2 为返波区。"
        Text 60, 496, 990, 14, "4. 删除 Kc_RefPos_x/y/z、Macro_Modex 或 Macro_SweepWatch_Enable 中任一参数即可重新弹出此对话框。"

        PushButton    20, 526, 90,  42, "About", .about
        CancelButton 120, 526, 90,  42, .btnCancel
        PushButton   870, 526, 90,  42, "Help", .help
        OKButton     970, 526, 90,  42

        Text 468, 540, 144, 14,"CST 宏程序参数设置"

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

    ' 循环直到用户确认; nResult=1 表示控件事件需要重新取值
    Do While Not bConfirmed
        Do
            nResult = Dialog(dlg)
            If nResult = 0 Then
                ShowParamsDialog = False
                Exit Function
            End If
        Loop While nResult = 1

        ' 收集模式选择: 全部 / 勾选框 / 文本框输入
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
            g_iPowerFlowSrc = PW_SRC_NATIVE
        Else
            g_iPowerFlowSrc = PW_SRC_INDIRECT
        End If

        Select Case dlg.Group3
            Case REGION_FORWARD
                g_iRegionMode = REGION_FORWARD
            Case REGION_BACKWARD
                g_iRegionMode = REGION_BACKWARD
            Case Else
                g_iRegionMode = REGION_BOTH
        End Select

        ' 周期方向坐标强制归零 (界面已置灰)
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
            ' 未选任何模式: 提示后重新显示对话框
            If dlg.Group1 = 2 Then
                MsgBox "请至少输入一个模式!", vbExclamation, "温馨提示"
            Else
                MsgBox "请至少勾选一个模式!", vbExclamation, "温馨提示"
            End If
        Else
            If g_bAllModes Then
                g_nSelectedModes = 0
            Else
                ' 自定义模式串: 统一分隔符后逐个校验 (非法/越界/重复一律忽略)
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
                        LogInfo "提示: 自定义模式输入 """ & sItem & """ 不是数字, 已忽略"
                    Else
                        dModeTmp = SafeCDbl(sItem, 0)
                        If dModeTmp <> Int(dModeTmp) Then
                            LogInfo "提示: 自定义模式输入 """ & sItem & """ 不是整数, 已忽略"
                        ElseIf dModeTmp < 1# Or dModeTmp > CDbl(g_iNumModes) Then
                            LogInfo "提示: 模式 " & sItem & " 超出求解器模式数范围 (1 - " & _
                                g_iNumModes & "), 已忽略"
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
                        LogInfo "提示: 输入的模式号含重复项, 已自动去重 (" & _
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
            sConfirmMsg = "请确认以下参数设置: " & vbCrLf & vbCrLf
            If g_bAllModes Then
                sConfirmMsg = sConfirmMsg & "计算模式: 所有模式 (1 - " & g_iNumModes & ")" & vbCrLf & vbCrLf
            Else
                sConfirmMsg = sConfirmMsg & "计算模式: " & selModes & vbCrLf & vbCrLf
            End If
            sConfirmMsg = sConfirmMsg & "参考位置 X: " & pos(1) & " " & g_sUnit & vbCrLf
            sConfirmMsg = sConfirmMsg & "参考位置 Y: " & pos(2) & " " & g_sUnit & vbCrLf
            sConfirmMsg = sConfirmMsg & "参考位置 Z: " & pos(3) & " " & g_sUnit & vbCrLf  & vbCrLf
            sConfirmMsg = sConfirmMsg & "功率流计算: 使用"& PowerFlowSrcText(g_iPowerFlowSrc) & vbCrLf
            sConfirmMsg = sConfirmMsg & "计算区域: " & RegionText(g_iRegionMode) & vbCrLf & vbCrLf
            sConfirmMsg = sConfirmMsg & "确认继续吗?"

            ' 二次确认: 未确认则回到对话框继续修改
            If MsgBox(sConfirmMsg, vbYesNo + vbQuestion, "确认参数设置") = vbYes Then
                bConfirmed = True
            End If
        End If
    Loop

    bAllModes = g_bAllModes
    ShowParamsDialog = True
End Function

' 关于对话框
Private Sub ShowAboutDialog()

    Begin Dialog UserDialog 1160, 370, "CST 本征模慢波结构用户监视器宏代码程序 — 关于本程序 版本 " & g_sVersionString & " (" & g_sVersionLabel & ")" 

        Text 416, 20, 328, 14, "相速归一化曲线 / 耦合阻抗 / 布里渊图 计算"
        Text  48, 42, 660, 14, "[ 代码编写者 ]  " & g_sAuthorName
        Text  48, 64, 1000, 14, "[ 编写时间点 ]  " & g_sReleaseDate & "      适用平台: " & g_sPlatform

        GroupBox 20, 86, 1120, 87, "[ 声明 ]"

        Text 48, 107, 1072, 14, "1. 编写者对宏程序的技术实现负责。"
        Text 48, 129, 1072, 14, "2. 本程序以 MIT 许可证开源发布, 允许自由使用、修改、商用与再分发, 需保留原始版权与许可声明 (详见随附 LICENSE)。"
        Text 48, 151, 1072, 14, "3. 代码按""原样""提供, 不保证计算结果的绝对准确性, 使用者应自行对关键仿真结果进行合理性验证；代码编写者不承担因此产生的任何责任。"

        GroupBox 20, 181, 1120, 131, "[ 运行前提 ]"

        Text 48, 202, 1072, 14, "1. 使用四面体或六面体网格算法, 推荐使用六面体网格 JDM 算法, 因为计算精确。"
        Text 48, 224, 1072, 14, "2. 慢波结构已设置周期边界条件。"
        Text 48, 246, 1072, 14, "3. 参数扫描列表中需包含 phase 参数。"
        Text 48, 268, 1072, 14, "4. 在 CST 宏命令中添加了 Slow Wave Userdefined Watch。"
        Text 48, 290, 1072, 14, "5. CST 版本号大于或等于 " & g_sCstVerMin & " 且不大于 " & g_sCstVerMax & ", 且正确部署核心计算库 LinXi.dll。"

        OKButton 530, 320, 100, 42

    End Dialog

    Dim dlg As UserDialog
    Dialog dlg

End Sub

' 使用帮助对话框 (参数说明、日志位置与重要注意事项)
Private Sub ShowHelpDialog()

    Begin Dialog UserDialog 1160, 645, "CST 本征模慢波结构用户监视器宏代码程序 — 使用帮助 版本 " & g_sVersionString & " (" & g_sVersionLabel & ")"

        Text 428, 20, 264, 14, "CST 本征模扫参宏程序 使用帮助手册"

        GroupBox 20, 42, 520, 295, "[ 参数说明 ]"

        Text  36,  63, 484, 14, "1. Macro_SweepWatch_Enable = 运行使能位 + 功率流计算方式 + 计算区域"
        Text  56,  91, 464, 14, "- 整数位: 1 = 间接功率流法; 2 = 原生功率流法; 0 = 禁用。"
        Text  56, 119, 464, 14, "- 小数位: 0 = 行波区与返波区; 1 = 仅行波区; 2 = 仅返波区。"
        Text  56, 147, 464, 14, "- 行波区要求频率随 phase 上升, 返波区要求频率随 phase 下降。"
        Text  56, 175, 464, 14, "- 其它取值 (例如 1.3、1.11) 会在下次运行时重新弹出参数设置对话框。"
        Text  36, 203, 484, 14, "2. Kc_RefPos_x / y / z = 耦合阻抗参考位置的三维坐标"
        Text  56, 231, 464, 14, "- 周期方向自动锁定为计算域中心, 不可修改。"
        Text  56, 259, 464, 14, "- 非周期方向可由用户自由设置。"
        Text  36, 287, 484, 14, "3. Macro_Modex (x = 1, 2, 3 ...) = 耦合阻抗计算标志"
        Text  56, 315, 464, 14, "- 宏仅检测该参数是否存在于参数列表中, 不关心其取值。"

        GroupBox 560, 42, 580, 295, "[ 参数设置与日志系统 ]"

        Text 576,  63, 548, 14, "-- 参数设置"
        Text 576,  91, 548, 14, "- 按界面提示勾选或输入模式编号即可。"
        Text 576, 119, 548, 14, "- 若设置的模式号超出求解器本征模数量, 将自动截断超过范围的部分。"
        Text 576, 147, 548, 14, "- 功率流的计算有两种方式可以选择。"
        Text 576, 175, 548, 14, "- 计算区域可选择行波区、返波区或两者都计算。"
        Text 576, 203, 548, 14, "- 选择两者时不再判断频率的升降, 全部扫描点都计算耦合阻抗。"
        Text 576, 231, 548, 14, "-- 日志系统"
        Text 576, 259, 548, 14, "- 日志文件: 工程目录 \Temp\macro_log.txt"
        Text 576, 287, 548, 14, "- 记录扫参全过程 (进度、错误信息), 方便排查问题。"

        GroupBox 20, 347, 1120, 239, "[ 重要注意事项 ]"

        Text  36, 368, 528, 14, "- 求解器推荐使用 JDM 算法, 因为计算精确。"
        Text 592, 368, 528, 14, "- JDM 算法下本征模命名不得含小数点。"
        Text  36, 396, 528, 14, "- 扫参期间请勿在当前窗口手动操作。"
        Text 592, 396, 528, 14, "- 重复扫参前请清除已有结果, 否则计算结果会丢失。"
        Text  36, 424, 528, 14, "- 单窗口单任务, 多任务请用多个 CST 窗口。"
        Text 592, 424, 528, 14, "- 暂停方式: 取消扫参任务即可。"
        Text  36, 452, 1088, 14, "- 删除历史树慢波结构用户监视器步骤不等于移除宏程序本身。"
        Text  36, 480, 1088, 14, "- 请勿开启本征模工程下电场或磁场的 Fields on Plane 和 Cutting Plane 视图。"
        Text  36, 508, 1088, 14, "- 宏程序所依赖的 Macro_SweepWatch_Enable, Macro_Modex 以及周期方向的 Kc_RefPos 参数不得扫描。"
        Text  36, 536, 1088, 14, "- 删除 Macro_SweepWatch_Enable 可让程序在下一次运行时重新弹出参数设置对话框, 重选功率流方式与计算区域。"
        Text  36, 564, 1088, 14, "- 建议将工程存放于路径较短的目录中, 避免路径长度超出 Windows 260 字符限制, 造成数据文件读写失败。"

        OKButton 530, 595, 100, 42

    End Dialog

    Dim dlg As UserDialog
    Dialog dlg

End Sub

' 安全字符串转 Double, 失败时返回给定默认值
Private Function SafeCDbl(ByVal sVal As String, ByVal dDefault As Double) As Double
    Err.Clear
    On Error Resume Next
    SafeCDbl = CDbl(Trim(sVal))
    If Err <> 0 Then SafeCDbl = dDefault
    Err.Clear
    On Error GoTo 0
End Function

' 功率流方式的中文名称, 用于界面与确认信息
Private Function PowerFlowSrcText(ByVal iSrc As Integer) As String
    If iSrc = 1 Then
        PowerFlowSrcText = "原生功率流法"
    Else
        PowerFlowSrcText = "间接功率流法"
    End If
End Function

' 计算区域的中文名称, 用于界面与确认信息
Private Function RegionText(ByVal iRegion As Integer) As String
    Select Case iRegion
        Case REGION_FORWARD
            RegionText = "行波区"
        Case REGION_BACKWARD
            RegionText = "返波区"
        Case Else
            RegionText = "行波区与返波区"
    End Select
End Function

' 使能位的完整中文说明, 用于参数描述与日志
Private Function EnableFlagText(ByVal iSrc As Integer, ByVal iRegion As Integer) As String
    EnableFlagText = "宏激活状态, 使用" & PowerFlowSrcText(iSrc) & _
        ", 计算" & RegionText(iRegion)
End Function

' 解析使能位: 整数位 = 功率流方式 (1 间接 / 2 原生), 一位小数位 = 计算区域 (0 双区 / 1 行波区 / 2 返波区)
' 0 (PW_SRC_INVALID) 由调用方按禁用处理; 其余不符合该规则的取值一律返回 False
Private Function DecodeEnableFlag(ByVal dValue As Double, _
    ByRef iSrc As Integer, ByRef iRegion As Integer) As Boolean

    Dim dScaled As Double
    Dim nCode As Long
    Dim nMethod As Long
    Dim nRegionCode As Long
    Dim nSrcCode As Long

    DecodeEnableFlag = False
    iSrc = 0
    iRegion = REGION_BOTH

    ' 非正值不在此处理(0 由调用方按禁用处理); 放大 10 倍后必须落在整数点上, 即小数位不超过一位
    If dValue <= 0# Then Exit Function
    dScaled = dValue * CDbl(ENABLE_FLAG_SCALE)
    If dScaled > 1000000# Then Exit Function
    ' CLng 采用银行家舍入, 故先用 Int 截断 (dScaled > 0, Int 即向下取整) 再转换
    nCode = CLng(Int(dScaled + 0.5))
    If Abs(dScaled - CDbl(nCode)) > ENABLE_FLAG_EPS Then Exit Function

    ' 先验证整数位与小数位的范围, 再收敛为 Integer, 避免非法大数值触发溢出
    nMethod = nCode \ ENABLE_FLAG_SCALE
    nRegionCode = nCode Mod ENABLE_FLAG_SCALE

    ' 合法整数位应为 PW_SRC_FROM_EH 或 PW_SRC_FROM_CST; 内部约定的位源编号 = 使能位整数位 - 1
    nSrcCode = nMethod - 1
    If nSrcCode <> PW_SRC_INDIRECT And nSrcCode <> PW_SRC_NATIVE Then Exit Function
    If nRegionCode < REGION_BOTH Or nRegionCode > REGION_BACKWARD Then Exit Function

    iSrc = CInt(nSrcCode)
    iRegion = CInt(nRegionCode)
    DecodeEnableFlag = True
End Function

' 判断字符串是否含 Windows 文件名非法字符
Private Function HasIllegalChar(ByVal s As String) As Boolean
    HasIllegalChar = (InStr(s, "\") > 0 Or InStr(s, "/") > 0 Or InStr(s, ":") > 0 Or _
        InStr(s, "*") > 0 Or InStr(s, "?") > 0 Or InStr(s, """") > 0 Or _
        InStr(s, "<") > 0 Or InStr(s, ">") > 0 Or InStr(s, "|") > 0)
End Function

' 把参数组合键改写成可用作文件名的形式 (替换非法字符、压缩下划线、去掉尾部空格与点)
Private Function SafeNamePart(ByVal s As String) As String
    Dim i As Long, ch As String, sOut As String

    ' 参数组合键由 " & " 与 " = " 拼接而成(等号后为两个空格), 替换串必须逐字对应
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

' 生成文件名标签; 超长或含非法字符时截断并追加分组序号 _gN
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

    ' 预算不足时用最小标记代替, 保证返回的标签长度不超过预算
    If nBudget < MIN_FILE_TAG_BUDGET Then
        MakeFileTag = "_g" & CStr(nIndex)
        Exit Function
    End If

    sRead = SafeNamePart(sKey)
    If Len(sRead) > nBudget - 6 Then sRead = Left$(sRead, nBudget - 6)

    MakeFileTag = sRead & "_g" & CStr(nIndex)
End Function

' 在 _groups.txt 中按组合键查序号, 顺便返回最大序号 nCount
Private Function LookupGroupIndex(ByVal sKey As String, ByRef nCount As Long) As Long
    Dim f As Integer
    Dim sLine As String, sK As String, sIdx As String
    Dim p1 As Long, p2 As Long, nIdx As Long

    nCount = 0
    LookupGroupIndex = 0
    If Len(sKey) = 0 Then Exit Function
    If Not FileExists(g_sTemp & "_groups.txt") Then Exit Function

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

' 根据 g_bAllModes 或选定列表生成各模式的计算标志表
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

' 递归删除结果树节点及其全部子节点
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

' 文件是否真实存在
Private Function FileExists(ByVal sPath As String) As Boolean
    Dim nAttr As Long

    FileExists = False
    If Len(sPath) = 0 Then Exit Function

    Err.Clear
    On Error Resume Next
    nAttr = GetAttr(sPath)
    If Err.Number = 0 Then FileExists = True
    Err.Clear
    On Error GoTo 0
End Function

' 按通配符批量删除文件 (删除失败静默忽略)
Private Sub DeleteFilesByPattern(ByVal sPattern As String)
    Dim sDir As String, sFile As String
    Dim nCut As Long

    ' 目标目录 = 模式串最后一个反斜杠及其之前的部分; 不含反斜杠时无法定位, 直接放弃
    nCut = InStrRev(sPattern, "\")
    If nCut = 0 Then Exit Sub
    sDir = Left$(sPattern, nCut)

    ' 本过程必须独占 Dir 的枚举状态: 循环体内不得调用 Dir 或任何会枚举目录的接口
    On Error Resume Next
    sFile = Dir(sPattern)
    Do While sFile <> ""
        Kill sDir & sFile
        sFile = Dir()
    Loop
    On Error GoTo 0
End Sub

' 模式编号冒泡升序排序
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

' 截断超出求解器模式数的编号; 全部越界时退回"计算全部模式"
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
        LogInfo "警告: 选择的模式全部超出范围(最大Mode = " & nMaxMode & "), 自动切换为所有模式"
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
        LogInfo "警告: 原始选择中包含超出范围模式号(最大Mode = " & nMaxMode & "), 已自动截断"
    End If
End Sub

' 模式编号是否在选定列表中
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

' 选中指定结果树条目并等待界面刷新 (读场前必须调用)
Private Sub SwitchToItem(ByVal sPath As String)
    On Error Resume Next
    SelectTreeItem sPath
    On Error GoTo 0
    DoEvents
    Wait WAIT_SEC
End Sub

' 设置标题与坐标轴标签后保存 1D 结果, 并挂到结果树
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

' 采样网格构建失败的统一处理: 清零网格参数并置失败标志
Private Sub FailGrid(ByVal sMsg As String)
    LogError sMsg
    g_n1 = 0
    g_n2 = 0
    g_nL = 0
    g_nCross = 0
    g_bGridFailed = True
End Sub

' 按网格类型生成采样网格与积分权重 (四面体走均匀网格, 六面体沿用 CST 网格); 每个参数组合只算一次
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
        FailGrid "功率流截面: 网格数为 0 (n1=" & g_n1 & ", n2=" & g_n2 & ")"
        Exit Sub
    End If

    Select Case g_iDir
        Case 1: iwLow = Mesh.GetClosestXIndex(CStr(g_cMin(1))): iwHigh = Mesh.GetClosestXIndex(CStr(g_cMax(1)))
        Case 2: iwLow = Mesh.GetClosestYIndex(CStr(g_cMin(2))): iwHigh = Mesh.GetClosestYIndex(CStr(g_cMax(2)))
        Case 3: iwLow = Mesh.GetClosestZIndex(CStr(g_cMin(3))): iwHigh = Mesh.GetClosestZIndex(CStr(g_cMax(3)))
    End Select

    g_nL = iwHigh - iwLow
    If g_nL <= 0 Then
        FailGrid "纵向积分: 网格数为 0"
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
        FailGrid "六面体采样 DLL 生成失败 (返回码 " & nRet & ")"
        Exit Sub
    End If

    Select Case g_iDir
        Case 1
            LogInfo "  采样网格分布: 周期方向 X = " & g_nL & _
                ", 横向方向 Y = " & g_n1 & ", 横向方向 Z = " & g_n2
        Case 2
            LogInfo "  采样网格分布: 周期方向 Y = " & g_nL & _
                ", 横向方向 X = " & g_n1 & ", 横向方向 Z = " & g_n2
        Case 3
            LogInfo "  采样网格分布: 周期方向 Z = " & g_nL & _
                ", 横向方向 X = " & g_n1 & ", 横向方向 Y = " & g_n2
    End Select

End Sub

' 四面体网格: 由计算域尺寸与网格边长求采样步长与点数
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
        FailGrid "四面体步长 DLL 计算失败 (返回码 " & nRet & ")"
        Exit Sub
    End If
    If n1 < 1 Or n2 < 1 Or nL < 1 Then
        FailGrid "四面体步长 DLL 返回无效维度 (n1=" & n1 & ", n2=" & n2 & ", nL=" & nL & ")"
        Exit Sub
    End If

    g_n1 = n1: g_n2 = n2: g_nL = nL
    LogInfo "  四面体自动步长: h = " & Format(h, "0.000E+00") & " " & g_sUnit
End Sub

' 四面体网格: 生成均匀采样网格、积分权重与采样点文件
Private Sub BuildTetraGrid(ByVal n1 As Long, ByVal n2 As Long, ByVal nL As Long)
    g_n1 = n1: g_n2 = n2: g_nL = nL
    If g_n1 < 1 Or g_n2 < 1 Or g_nL < 1 Then
        FailGrid "四面体采样网格维度无效 (n1=" & g_n1 & ", n2=" & g_n2 & _
            ", nL=" & g_nL & "), 跳过本次计算"
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
        FailGrid "四面体网格几何 DLL 生成失败 (返回码 " & nRet & ")"
        Exit Sub
    End If
End Sub

' 四面体网格的采样网格构建流程 (定步长 -> 建网格)
Private Sub PrecomputeRegularGrid()
    DetermineTetraStep
    If g_n1 < 1 Or g_n2 < 1 Or g_nL < 1 Then Exit Sub
    BuildTetraGrid g_n1, g_n2, g_nL
    If g_nCross < 1 Then Exit Sub
    Select Case g_iDir
        Case 1
            LogInfo "  采样网格分布: 周期方向 X = " & g_nL & _
                ", 横向方向 Y = " & g_n1 & ", 横向方向 Z = " & g_n2
        Case 2
            LogInfo "  采样网格分布: 周期方向 Y = " & g_nL & _
                ", 横向方向 X = " & g_n1 & ", 横向方向 Z = " & g_n2
        Case 3
            LogInfo "  采样网格分布: 周期方向 Z = " & g_nL & _
                ", 横向方向 X = " & g_n1 & ", 横向方向 Y = " & g_n2
    End Select

End Sub

' 把某个模式的指定场量导出到采样点文件 (ASCIIExport)
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

' 调用 DLL 解析场文件为 Ex/Ey/Ez (实部与虚部), 校验点数
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
        LogError "解析场文件失败: " & sPath & " (返回码 " & nRet & ")"
        ParseFieldFile = False
    End If
End Function

' 调用 DLL 解析功率流场文件为 Px/Py/Pz (实部与虚部)
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
        LogError "解析功率流场文件失败: " & sPath & " (返回码 " & nRet & ")"
        ParsePowerFlowFile = False
    End If
End Function

' 导出并解析 E 场
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

' 导出并解析 CST 原生功率流场
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

' 导出并解析 H 场
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

' 直接由结果对象读取本征模频率, 失败返回 0 以便回退到界面读取
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
        LogWarning "  Mode " & sMode & ": 频率读取失败, 切换获取方式"
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
        LogWarning "  Mode " & sMode & ": 频率读取失败, 切换获取方式"
        Exit Function
    End If

    R3DGetFieldFrequency = dFreq
End Function

' 间接法取场: 依次取得 E 场与 H 场
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

' 原生法取场: 由 E 与 H* 叉积生成功率流矢量场, 保存并加入结果树后导出
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
        LogError "  Mode " & sMode & ": 电磁场数据文件加载失败 (Err=" & nErr & ")"
    ElseIf resEField Is Nothing Or resHField Is Nothing Then
        LogError "  Mode " & sMode & ": 功率流矢量场生成失败, 未找到电磁场数据文件"
    ElseIf resEField.GetLength <= 0 Or resEField.GetLength <> resHField.GetLength Then
        LogError "  Mode " & sMode & ": 功率流矢量场生成失败, 可能是电磁场数据文件加载失败或文件顺序不一致引起的"
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
            LogError "  Mode " & sMode & ": 功率流矢量场生成 / 保存失败 (Err=" & nErr & ")"
        Else
            LogInfo "  Mode " & sMode & ": 已成功生成功率流矢量场"

            If ExportPAndParse(iMode) Then
                PreparePowerFieldFromCST = True
            Else
                LogError "  Mode " & sMode & ": 功率流场导出 / 解析失败"
            End If
        End If
    End If

    Set resPower = Nothing
    Set resHConj = Nothing
    Set resEField = Nothing
    Set resHField = Nothing
End Function

' 沿周期方向积分得到等效电压 V; 失败时输出置零
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
        LogError "纵向积分 DLL 调用失败 (返回码 " & nRet & ")"
        dVre = 0#: dVim = 0#: dEabs = 0#
    End If
End Sub

' 原生法功率流积分 (对 CST 功率流场沿横截面加权)
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

' 间接法功率流积分 (由 E、H 场现场叉积后加权)
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

' 临时结果文件路径 (_<名称>_<模式>.sig)
Private Function TempPath(ByVal sBase As String, ByVal iMode As Integer) As String
    TempPath = g_sTemp & "_" & sBase & "_" & CStr(iMode) & ".sig"
End Function

' 创建空文件 (先删除同名旧文件)
Private Sub CreateEmptyFile(ByVal sPath As String)
    On Error Resume Next
    Kill sPath
    On Error GoTo 0
    Dim f As Integer
    f = FreeFile
    Open sPath For Output As #f
    Close #f
End Sub

' 向结果行缓冲追加一个采样点 (频率 + 数值, 高精度科学计数法)
Private Sub BufferAppend(ByRef sBuf As String, ParamArray vals() As Variant)
    Dim i As Integer
    For i = 0 To UBound(vals)
        sBuf = sBuf & Format(CDbl(vals(i)), "0.000000e+000   ")
    Next i
    sBuf = sBuf & vbCrLf
End Sub

' 把缓冲内容批量追加写入文件; 写失败记录日志并保留现场
Private Sub BufferFlush(ByVal sPath As String, ByRef sBuf As String)
    If Len(sBuf) = 0 Then Exit Sub
    Dim f As Integer
    Err.Clear
    On Error Resume Next
    f = FreeFile
    Open sPath For Append As #f
    If Err <> 0 Then
        LogError "临时结果文件打开失败 (Err=" & Err.Number & "): " & sPath
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Print #f, sBuf;
    If Err <> 0 Then
        LogError "临时结果文件写入失败 (Err=" & Err.Number & "): " & sPath
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

' 复位各模式的相位有效性状态
Private Sub ResetPhaseSweepState()
    Dim i As Long
    If g_nPhaseStateModes <= 0 Then Exit Sub
    For i = 1 To g_nPhaseStateModes
        g_abPhaseDead(i) = False
        g_anPhaseFlat(i) = 0
        g_abPhaseFlatWarned(i) = False
        g_abKcWarned(i) = False
    Next i
End Sub

' 是否还有未落盘的结果行缓冲
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

' 清空结果行缓冲
Private Sub ClearRowBuffers()
    Dim i As Long
    If g_nBufModes <= 0 Then Exit Sub
    For i = 1 To g_nBufModes
        g_aBufBeta(i) = ""
        g_aBufZpierce(i) = ""
        g_aBufPhase(i) = ""
    Next i
End Sub

' 初始化日志: 建文件、记录本次运行起始时间与路径预算, 并校验核心计算库
Private Sub InitLogFile()
    LogInit
    If Len(g_sLogFile) = 0 Then Exit Sub

    If Not FileExists(g_sLogFile) Then CreateEmptyFile g_sLogFile

    Dim fTime As Integer
    fTime = FreeFile
    Open g_sTemp & "_runtime_start.txt" For Output As #fTime
    Print #fTime, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #fTime

    LogInfo " ----- 新的参数扫描任务开始, 起始时间: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ----- "
    LogInfo " ----- 日志文件: " & g_sLogFile & " ----- "
    LogInfo " ----- 路径预算: Temp 路径长度 = " & Len(g_sTemp) & " 字符, 文件名标签可用 = " & g_nFileTagBudget & " 字符 ----- "

    If g_nFileTagBudget < 12 Then LogWarning "工程路径过深: 文件名标签预算仅 " & g_nFileTagBudget & " 字符, 建议将工程移至更短的目录, 否则计算数据可能写入失败"

    CheckCoreLibrary
End Sub

' 日志模块的惰性初始化 (幂等)
Private Sub LogInit()
    If g_bLogInited Then Exit Sub
    g_bLogInited = True
    g_nLogFileMinLevel = llInfo
    g_sLogLastError = ""
End Sub

' 日志超过 LOG_ROTATE_BYTES 时改名备份
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

' 写一行日志 (带时间戳与等级); 失败时记录错误并转入 Debug 输出
Private Function WriteLog(ByVal sLevel As String, ByVal sMsg As String) As Boolean
    Dim f As Integer
    Dim nErr As Long
    Dim sErr As String
    Dim sLine As String

    If Len(g_sLogFile) = 0 Then Exit Function

    If Len(sMsg) > LOG_MSG_MAX_LEN Then sMsg = Left$(sMsg, LOG_MSG_MAX_LEN) & " ...(已截断)"
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

    g_sLogLastError = "日志写入失败 (Err=" & nErr & "): " & sErr & " | 文件: " & g_sLogFile
    Debug.Print g_sLogLastError
End Function

' 按等级统一分发: 写日志文件 + 输出到 CST 消息窗口, 必要时中止宏
Private Sub LogMessage(ByVal nLevel As LogLevel, ByVal sMsg As String, _
    Optional ByVal bAbortAfter As Boolean = False)
    Dim sTag As String
    Dim sPfx As String
    Dim bToWindow As Boolean

    If Len(Trim$(sMsg)) = 0 Then sMsg = "(空消息)"

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
        LogCritical "发现致命错误! 宏程序已中止扫参程序进行! 请检查各项参数设置后再重新开始扫参程序或执行宏代码程序!"
        ReportError MSG_ABORT_MACRO & sMsg & vbCrLf
    End If
End Sub

' 输出 DEBUG 级日志 (只进文件)
Private Sub LogDebug(ByVal sMsg As String)
    LogMessage llDebug, sMsg
End Sub

' 输出 INFO 级日志
Private Sub LogInfo(ByVal sMsg As String)
    LogMessage llInfo, sMsg
End Sub

' 输出 WARN 级日志
Private Sub LogWarning(ByVal sMsg As String)
    LogMessage llWarning, sMsg
End Sub

' 输出 ERROR 级日志
Private Sub LogError(ByVal sMsg As String)
    LogMessage llError, sMsg
End Sub

' 输出 CRIT 级日志; bAbortAfter=True 时中止扫参
Private Sub LogCritical(ByVal sMsg As String, Optional ByVal bAbortAfter As Boolean = False)
    LogMessage llCritical, sMsg, bAbortAfter
End Sub

' CST 宏入口: 依次执行 初始化 / 逐点处理 / 结果汇总 三个阶段 (便于在宏菜单中直接运行)
Public Sub Main()
    ParameterSweepWatch 0
    ParameterSweepWatch 1
    ParameterSweepWatch 2
End Sub
