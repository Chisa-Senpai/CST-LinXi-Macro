'==============================================================================
' CST-LinXi-Macro  -  version check (She.bas)
'
' Copyright (c) 2026 Limorazp
' SPDX-License-Identifier: MIT
'
' Released under the MIT License; full text: see LICENSE in the repository root.
'
'------------------------------------------------------------------------------
' Deployed as: Check LinXi Version.mcr
'
'==============================================================================

Option Explicit

Private Const CORE_AUTHOR   As String = " Lin Xi and She & Me "
Private Const DLL_FILE_NAME As String = "LinXi.dll"

Declare Function CCoreVersion Lib "LinXi.dll" (ByVal author As String) As Long

Private Const CFG_FILE_NAME   As String = "LinXi.ini"
Private Const CFG_SECTION_VER As String = "VERSION"

Private Const DLL_OK        As Integer = 0
Private Const DLL_NOT_FOUND As Integer = 1
Private Const DLL_CALL_FAIL As Integer = 2
Private Const DLL_TOO_OLD   As Integer = 3

Private g_sVersionString As String
Private g_sVersionLabel  As String
Private g_nDllMinVer     As Long
Private g_sCfgPath       As String
Private g_sProblems      As String

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

Private Sub ShowVersionDialog(ByVal nDllVer As Long)
    Dim sCstVer As String

    sCstVer = CstVersionText()

    Begin Dialog UserDialog 440, 204, "LinXi Macro - Program version information"

        Text 78, 12, 290, 16, "CST slow-wave structure user watch macro"
        Text 128, 32, 200, 14, "VERSION INFORMATION"

        GroupBox 20, 56, 400, 88, ""

        Text  32,  70, 376, 14, "Macro version : " & g_sVersionString & "  (" & g_sVersionLabel & ")"
        Text  32,  92, 376, 14, "Core library : " & DLL_FILE_NAME & " v" & CStr(nDllVer) & "  (minimum required v" & CStr(g_nDllMinVer) & ")"
        Text  32, 114, 376, 14, "CST version : " & sCstVer

        OKButton 170, 158, 100, 42

    End Dialog

    Dim dlg As UserDialog
    Dialog dlg
End Sub

Private Sub ShowProblemDialog(ByVal bCfgOK As Boolean, ByVal iDllStatus As Integer, _
    ByVal nDllVer As Long, ByVal sDllNote As String)

    Dim sMsg As String

    sMsg = "The following problems were detected:" & vbCrLf & vbCrLf & g_sProblems & vbCrLf & vbCrLf & _
        "Information obtained:" & vbCrLf & _
        "  Macro version : " & VersionText(bCfgOK) & vbCrLf & _
        "  Core library  : " & DllText(iDllStatus, nDllVer, sDllNote) & vbCrLf & _
        "  CST version   : " & CstVersionText() & vbCrLf & vbCrLf & _
        "Note: re-running the installer restores the default deployment;" & vbCrLf & _
        "      if LinXi.ini was just edited, make sure it is saved as ANSI/GBK."

    MsgBox sMsg, vbCritical, "Version check - problems detected"
End Sub

Private Function VersionText(ByVal bOK As Boolean) As String
    If bOK Then
        VersionText = g_sVersionString & " (" & g_sVersionLabel & ")"
    Else
        VersionText = "(not available)"
    End If
End Function

Private Function DllText(ByVal iStatus As Integer, ByVal nVer As Long, ByVal sNote As String) As String
    If iStatus = DLL_OK Then
        DllText = DLL_FILE_NAME & " v" & CStr(nVer) & " (minimum required v" & CStr(g_nDllMinVer) & ")"
    ElseIf nVer > 0 Then
        DllText = DLL_FILE_NAME & " v" & CStr(nVer) & " -- " & sNote
    ElseIf Len(sNote) > 0 Then
        DllText = DLL_FILE_NAME & " (" & sNote & ")"
    Else
        DllText = "(not available)"
    End If
End Function

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
        AddProblem "None of the 4 candidate locations contains the external configuration file " & CFG_FILE_NAME & _
            ", so the current macro version cannot be determined." & vbCrLf & _
            "    The macro searched for it in this order:" & vbCrLf & CfgCandidateList()
        Exit Function
    End If

    If Not CfgRead(g_sCfgPath, sVer, sLbl, sDllMin) Then
        AddProblem "The external configuration file cannot be read (in use or insufficient permissions):" & vbCrLf & _
            "    " & g_sCfgPath
        Exit Function
    End If

    If Len(sVer) = 0 Or Len(sLbl) = 0 Or Len(sDllMin) = 0 Then
        AddProblem "The [Version] section of the configuration file is missing the mandatory keys VersionString / VersionLabel / DllMinVersion:" & vbCrLf & _
            "    " & g_sCfgPath & vbCrLf & _
            "    If the file was just edited, make sure it is saved as ANSI/GBK -- saving it as UTF-8 makes every key unreadable."
        Exit Function
    End If

    If Not IsNumeric(sDllMin) Then
        AddProblem "DllMinVersion in the configuration file is not a valid version number: """ & sDllMin & _
            """ (it must be a positive integer, for example 122)."
        Exit Function
    End If

    g_nDllMinVer = CLng(Val(sDllMin))
    If g_nDllMinVer <= 0 Then
        AddProblem "DllMinVersion in the configuration file must be a positive integer, but it is """ & sDllMin & """."
        Exit Function
    End If

    g_sVersionString = sVer
    g_sVersionLabel = sLbl
    LoadVersionFromIni = True
End Function

Private Sub AddProblem(ByVal sMsg As String)
    If Len(g_sProblems) > 0 Then g_sProblems = g_sProblems & vbCrLf & vbCrLf
    g_sProblems = g_sProblems & "- " & sMsg
End Sub

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
        sNote = "cannot resolve the file path"
        AddProblem "Cannot resolve the path of the core library " & DLL_FILE_NAME & ": the GetInstallPath call failed."
        Exit Function
    End If

    If Not DiskFileExists(sDllPath) Then
        QueryDllVersion = DLL_NOT_FOUND
        sNote = "file missing"
        AddProblem "Core library file not found:" & vbCrLf & _
            "    " & sDllPath & vbCrLf & _
            "    Re-run the installer to deploy it."
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
        sNote = "call failed"
        AddProblem "The core library file exists, but the CCoreVersion call failed:" & vbCrLf & _
            "    Err.Number   = " & CStr(nErrNum) & vbCrLf & _
            "    LastDLLError = " & CStr(nDllErr) & vbCrLf & _
            "    Description  = " & sErrDesc & vbCrLf & _
            "    Common causes: the DLL is 32-bit (it must be 64-bit) / exported name mismatch / missing dependency."
        Exit Function
    End If

    If nVer < g_nDllMinVer Then
        QueryDllVersion = DLL_TOO_OLD
        sNote = "version too old"
        AddProblem "The core library is too old: currently v" & CStr(nVer) & _
            ", but this program requires at least v" & CStr(g_nDllMinVer) & vbCrLf & _
            "    (that requirement is written in DllMinVersion of " & CFG_FILE_NAME & ")"
        Exit Function
    End If
End Function

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
    If Len(sVer) = 0 Then sVer = "(unavailable)"
    CstVersionText = sVer
End Function
