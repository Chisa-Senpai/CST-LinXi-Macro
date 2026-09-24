'==============================================================================
' CST-LinXi-Macro  -  slow-wave structure watch setup wizard (Me.bas)
'
' Copyright (c) 2026 Limorazp
' SPDX-License-Identifier: MIT
'
' Released under the MIT License; full text: see LICENSE in the repository root.
'
'------------------------------------------------------------------------------
' Deployed as: Define LinXi Macro.mcr
'
'==============================================================================

Option Explicit

Private Const WATCH_FILE_NAME As String = "LinXi_Watch.bas"

Private Const WATCH_LEGACY_DIR As String = "\Library\Macros\Solver\E-Solver\"
Private Const WATCH_LEGACY_NAME As String = "Model.pfc"

Private Const PROJECT_MACRO_NAME As String = "Model.pfc"

Private Const CFG_FILE_NAME   As String = "LinXi.ini"
Private Const CFG_SECTION_VER As String = "VERSION"

Private g_sVersionString As String
Private g_sVersionLabel  As String
Private g_sAuthorName    As String
Private g_sCfgPath       As String
Private g_sCfgNote       As String

Sub Main()
    Dim bWatchExists As Boolean
    Dim bRegistered  As Boolean
    Dim sWatchMacro  As String

    If Not LoadCfgForWizard() Then Exit Sub

    sWatchMacro = WatchMacroFind()
    If Len(sWatchMacro) = 0 Then
        ShowWatchMissingDialog
        Exit Sub
    End If

    If Not ShowWelcomeDialog() Then
        MsgBox "The user cancelled adding the SWS user watch.", vbInformation, "Cancelled"
        Exit Sub
    End If

    bWatchExists = CopyWatchMacro(sWatchMacro)

    If Not bWatchExists Then
        MsgBox "Failed to deploy the watch macro into the project directory:" & vbCrLf & vbCrLf & _
            ProjectMacroPath() & vbCrLf & vbCrLf & _
            "Check that the project directory is writable, or re-run the installer.", _
            vbCritical, "Add failed"
        Exit Sub
    End If

    bRegistered = AddWatchToHistory()

    If bRegistered Then
        MsgBox "The SWS user watch was added successfully!" & vbCrLf & vbCrLf & _
            "Watch macro: " & sWatchMacro & vbCrLf & _
            "Project file: " & ProjectMacroPath() & vbCrLf & vbCrLf & _
            "Just start the parameter sweep task, no macro edit is needed!", vbInformation, "Done"
    Else
        MsgBox "The watch macro was copied into the project, but registering it in the project history failed." & vbCrLf & vbCrLf & _
            "Please run this macro (Define LinXi Macro) again.", vbExclamation, "Partly done"
    End If
End Sub

Private Function AddWatchToHistory() As Boolean
    AddWatchToHistory = False

    On Error Resume Next
    AddToHistory "Add Watch: LinXi Macro", "ParameterSweep.AddUserdefinedWatch"
    If Err.Number = 0 Then AddWatchToHistory = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function ShowWelcomeDialog() As Boolean

    Begin Dialog UserDialog 640, 395, "Slow-Wave Structure User-Defined Watch Setup Wizard"

        Text 210, 20, 300, 14, "Welcome to the watch setup wizard!"

        GroupBox 20, 42, 600, 71, "[ Function ]"

        Text 40, 63, 560, 14, "Computes phase velocity and coupling impedance results."
        Text 40, 91, 560, 14, "Produces three curves: Brillouin, normalized vp and Kc."

        GroupBox 20, 121, 600, 127, "[ Notes ]"

        Text 40, 142, 560, 14, "1. The watch only needs to be configured once per project file."
        Text 40, 170, 560, 14, "2. Removing the watch step does not remove the macro itself."
        Text 40, 198, 560, 14, "3. Make sure LinXi_Watch.bas and LinXi.dll are deployed in CST's AMD64 folder."
        Text 40, 226, 560, 14, "4. Do not operate the CST window during a sweep; use one task."

        GroupBox 20, 256, 600, 43, "[ Important Warning ]"

        Text 40, 277, 560, 14, "Adding the watch deletes the existing result data. Continue?"

        CancelButton 200, 320, 100, 42
        OKButton     340, 320, 100, 42

        Text  20, 372, 300, 14, "Version: " & g_sVersionString & " (" & g_sVersionLabel & ")"
        Text 512, 372, 200, 14, "Author:  " & g_sAuthorName

    End Dialog

    Dim dlg As UserDialog
    If Dialog(dlg) Then
        ShowWelcomeDialog = True
    Else
        ShowWelcomeDialog = False
    End If
End Function

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

Private Sub ShowWatchMissingDialog()
    Dim sMsg As String
    Dim i As Integer

    sMsg = "Error: the watch macro " & WATCH_FILE_NAME & " was not found, the operation was aborted." & vbCrLf & vbCrLf & _
        "The macro searched for it in this order, so at least one must exist:" & vbCrLf
    For i = 1 To 2
        sMsg = sMsg & "  " & i & ". " & WatchMacroCandidate(i) & vbCrLf
    Next i
    sMsg = sMsg & vbCrLf & _
        "Please re-run the installer: it deploys " & WATCH_FILE_NAME & _
        " together with " & CFG_FILE_NAME & " into the AMD64 folder of the CST installation."

    MsgBox sMsg, vbCritical, "Watch macro not found"
End Sub

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
        g_sCfgNote = "None of the 4 candidate locations contains the external configuration file " & CFG_FILE_NAME
    ElseIf Not CfgRead(g_sCfgPath, sVer, sLbl, sAuthor) Then
        g_sCfgNote = "The external configuration file cannot be read (in use or insufficient permissions): " & g_sCfgPath
    ElseIf Len(sVer) = 0 Or Len(sLbl) = 0 Or Len(sAuthor) = 0 Then
        g_sCfgNote = "The [Version] section of the configuration file is missing the mandatory keys VersionString / VersionLabel / Author: " & g_sCfgPath
    Else
        g_sVersionString = sVer
        g_sVersionLabel = sLbl
        g_sAuthorName = sAuthor
    End If

    If Len(g_sVersionString) = 0 Then
        MsgBox CfgErrorText(), vbCritical, "Macro aborted"
        Exit Function
    End If

    LoadCfgForWizard = True
End Function

Private Function CfgErrorText() As String
    Dim sMsg As String
    Dim i As Integer

    sMsg = "Error: the version information must come from the external configuration file " & CFG_FILE_NAME & _
        ", but it is currently unavailable, so adding the watch was aborted!" & vbCrLf & vbCrLf & _
        "Reason: " & g_sCfgNote & vbCrLf & vbCrLf & _
        "The macro searched for the file in this order, so at least one must exist and be readable:" & vbCrLf
    For i = 1 To 4
        If Len(CfgCandidatePath(i)) > 0 Then
            sMsg = sMsg & "  " & i & ". " & CfgCandidatePath(i) & vbCrLf
        End If
    Next i
    sMsg = sMsg & vbCrLf & _
        "Note: the file must be saved as ANSI/GBK; saving it as UTF-8 makes every key unreadable." & vbCrLf & _
        "      Re-run the installer to restore the default deployment."

    CfgErrorText = sMsg
End Function

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

Private Function CfgExists(ByVal sPath As String) As Boolean
    CfgExists = DiskFileExists(sPath)
End Function

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
