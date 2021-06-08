Attribute VB_Name = "modMain"
' //
' // modMain.bas - startup module with global functions
' // By The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME = "modMain"

Public Enum eSettings
    SET_THEME       ' // Current theme
    SET_FONT_SIZE   ' // Current font size
End Enum

Public Enum eTheme
    T_WIN
    T_DARK
    T_SOFT
End Enum

Public Enum eFontSize
    FS_SMALL
    FS_LARGE
End Enum

Public Sub Main()
    Const PROC_NAME = "Main", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim tICC    As tagINITCOMMONCONTROLSEX
    
    On Error GoTo error_handler
    
    tICC.dwSize = Len(tICC)
    tICC.dwICC = ICC_WIN95_CLASSES
    
    InitCommonControlsEx tICC

    frmMain.Show

    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

' // Show current error message
Public Sub ShowCurrentError()
    MsgBox "An error occured 0x" & Hex$(Err.Number) & vbNewLine & "Source: " & Err.Source & _
            vbNewLine & vbNewLine & Err.Description, vbCritical
End Sub

' // Throw error
Public Sub ThrowCurrentErrorUp( _
           ByRef sProcName As String)
    If StrComp(Err.Source, sProcName, vbTextCompare) = 0 Or Len(Err.Source) = 0 Then
        Err.Raise Err.Number, sProcName, Err.Description
    Else
        Err.Raise Err.Number, sProcName, Err.Source & vbNewLine & Err.Description
    End If
End Sub

' // Save settings
Public Property Let Settings( _
                    ByVal eParam As eSettings, _
                    ByVal lValue As Long)
    Const PROC_NAME = "Settings_put", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim sSection    As String
    Dim sKey        As String
    Dim sValue      As String
    
    On Error GoTo error_handler
    
    sSection = "Appearance"
    
    Select Case eParam
    Case SET_THEME
    
        sKey = "Theme"
        
        Select Case lValue
        Case T_DARK
            sValue = "Dark"
        Case T_SOFT
            sValue = "Soft"
        Case Else
            sValue = "Windows"
        End Select
        
    Case SET_FONT_SIZE
    
        sKey = "FontSize"
        
        Select Case lValue
        Case FS_LARGE
            sValue = "Large"
        Case Else
            sValue = "Small"
        End Select
        
    End Select
    
    If WritePrivateProfileString(sSection, sKey, sValue, App.Path & "\config.ini") = 0 Then
        Err.Raise 7, FULL_PROC_NAME, "WritePrivateProfileString failed"
    End If
    
    Exit Property
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Property

' // Load settings
Public Property Get Settings( _
                    ByVal eParam As eSettings) As Long
    Const PROC_NAME = "Settings_get", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim sSection    As String
    Dim sKey        As String
    Dim sRet        As String
    Dim lRet        As Long
    
    On Error GoTo error_handler
    
    sSection = "Appearance"
    
    Select Case eParam
    Case SET_THEME
        sKey = "Theme"
    Case SET_FONT_SIZE
        sKey = "FontSize"
    End Select
    
    sRet = Space$(255)
    
    lRet = GetPrivateProfileString(sSection, sKey, vbNullString, sRet, Len(sRet), App.Path & "\config.ini")
    
    If lRet Then
        sRet = Left$(sRet, lRet)
    Else
        sRet = vbNullString
    End If
    
    Select Case eParam
    Case SET_THEME
    
        If StrComp(sRet, "dark", vbTextCompare) = 0 Then
            Settings = T_DARK
        ElseIf StrComp(sRet, "soft", vbTextCompare) = 0 Then
            Settings = T_SOFT
        Else
            Settings = T_WIN
        End If
        
    Case SET_FONT_SIZE
    
        If StrComp(sRet, "large", vbTextCompare) = 0 Then
            Settings = FS_LARGE
        Else
            Settings = FS_SMALL
        End If
        
    End Select
    
    Exit Property
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Property

' // Convert UNICODE text to UNICODE with allowed characters
Public Function FixUnicode( _
                ByRef sValue As String) As String
    Const PROC_NAME = "FixUnicode", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bKOI()  As Byte
    
    On Error GoTo error_handler
    
    bKOI = UnicodeToVec6KOI7(sValue)
    FixUnicode = Vec6KOI7ToUnicode(bKOI)
    
    Exit Function
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Function

' // Convert token to keyword or symbol
Public Function Vec6BasicKeyword( _
                ByVal bCode As Byte) As String
    Const PROC_NAME = "Vec6BasicKeyword", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Static s_sTable()   As String
    Static s_bInit      As Boolean
    
    Dim bTemp() As Byte
    Dim lIndex  As Long
    Dim lIndex2 As Long
    Dim lSize   As Long
    
    On Error GoTo error_handler
    
    If Not s_bInit Then
        
        bTemp = LoadResData(100, RT_RCDATA)
        
        ReDim s_sTable(228)
            
        For lIndex = 0 To UBound(s_sTable)
                    
            lSize = lstrlenW(bTemp(lIndex2))
            
            If lSize = 0 Then
                lSize = 1
            End If
            
            s_sTable(lIndex) = Space$(lSize)
            memcpy ByVal StrPtr(s_sTable(lIndex)), bTemp(lIndex2), lSize * 2
            lIndex2 = lIndex2 + (lSize + 1) * 2
            
        Next
        
        s_bInit = True
        
    End If
    
    If bCode > 228 Then
        Err.Raise 5, FULL_PROC_NAME, "Invalid code"
    End If

    Vec6BasicKeyword = s_sTable(bCode)
    
    Exit Function
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Function

' // Convert native Vector-06C KOI-7 to UNICODE symbols
Public Function Vec6KOI7ToUnicode( _
                ByRef bValue() As Byte) As String
    Const PROC_NAME = "Vec6KOI7ToUnicode", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Static s_bTable()   As Byte
    
    Dim lCount  As Long
    Dim bOut()  As Byte
    Dim lIndex  As Long
    
    On Error GoTo error_handler
    
    If SACount(ArrPtr(s_bTable)) = 0 Then
        s_bTable = LoadResData(101, RT_RCDATA)
    End If
    
    lCount = SACount(ArrPtr(bValue))
    If lCount = 0 Then Exit Function
    
    ReDim bOut(lCount * 2 - 1)
    
    For lIndex = 0 To lCount - 1
        
        If bValue(lIndex) > 127 Then
            Err.Raise 5, FULL_PROC_NAME, "Invalid symbol"
        End If
        
        GetMem2 s_bTable(bValue(lIndex) * 2), bOut(lIndex * 2)
               
    Next
    
    Vec6KOI7ToUnicode = bOut
    
    Exit Function
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Function

' // Convert UNICODE string to native Vector-06C KOI-7
Public Function UnicodeToVec6KOI7( _
                ByRef sValue As String) As Byte()
    Const PROC_NAME = "UnicodeToVec6KOI7", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Static s_bTable()   As Byte
    
    Dim bRet()  As Byte
    Dim bIn()   As Byte
    Dim lIndex  As Long
    Dim lChar   As Long
    
    On Error GoTo error_handler
    
    If SACount(ArrPtr(s_bTable)) = 0 Then
        s_bTable = LoadResData(102, RT_RCDATA)
    End If
    
    If Len(sValue) = 0 Then Exit Function
    
    ReDim bRet(Len(sValue) - 1)
    
    bIn = sValue
    
    For lIndex = 0 To Len(sValue) - 1
        GetMem2 bIn(lIndex * 2), lChar
        bRet(lIndex) = s_bTable(lChar)
    Next
    
    UnicodeToVec6KOI7 = bRet
    
    Exit Function
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Function

' // Save byte array to file
' // Returns true if successful
Public Function SaveArrayToFile( _
                ByRef sFileName As String, _
                ByRef bData() As Byte) As Boolean
    Dim hFile       As Long
    Dim lSize       As Long
    Dim lWritten    As Long
    
    hFile = CreateFile(sFileName, GENERIC_READ Or GENERIC_WRITE, 0, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    
    If hFile = INVALID_HANDLE_VALUE Then
        Exit Function
    End If
    
    lSize = SACount(ArrPtr(bData))
    
    If lSize > 0 Then
    
        If WriteFile(hFile, bData(0), lSize, lWritten, ByVal 0&) = 0 Then
            GoTo CleanUp
        End If
        
        If lSize <> lWritten Then
            GoTo CleanUp
        End If
        
    End If
    
    SaveArrayToFile = True
    
CleanUp:
    
    CloseHandle hFile
    
End Function

' // Load byte array from file
' // Returns true if successful
' // lSize receives the size of loaded data
Public Function LoadFileToArray( _
                ByRef sFileName As String, _
                ByRef bData() As Byte, _
                ByRef lSize As Long) As Boolean
    Dim hFile   As Long
    Dim liSize  As LARGE_INTEGER
    Dim lRead   As Long
    
    hFile = CreateFile(sFileName, GENERIC_READ, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    If hFile = INVALID_HANDLE_VALUE Then
        Exit Function
    End If
    
    If GetFileSizeEx(hFile, liSize) = 0 Then
        GoTo CleanUp
    End If
    
    If liSize.HighPart <> 0 Or liSize.LowPart < 0 Or liSize.LowPart > &H6400000 Then
        GoTo CleanUp
    End If
    
    ReDim bData(liSize.LowPart - 1)
    
    If ReadFile(hFile, bData(0), liSize.LowPart, lRead, ByVal 0&) = 0 Then
        GoTo CleanUp
    End If
    
    If lRead <> liSize.LowPart Then
        GoTo CleanUp
    End If
    
    lSize = lRead
    
    LoadFileToArray = True
    
CleanUp:
    
    CloseHandle hFile
    
End Function

' // Get file extenstion
Public Function GetFileExtension( _
                ByRef sPath As String) As String
    Dim pExt    As Long
    
    pExt = PathFindExtension(ByVal StrPtr(sPath))
    
    GetFileExtension = Mid$(sPath, (pExt - StrPtr(sPath)) \ 2 + 1)
    
End Function

' // Get file title (only filename without extension)
Public Function GetFileTitle( _
                ByRef sPath As String) As String
    Dim pExt    As Long
    Dim pName   As Long
    
    pName = PathFindFileName(sPath)
    pExt = PathFindExtension(ByVal pName)
    
    GetFileTitle = Mid$(sPath, (pName - StrPtr(sPath)) \ 2 + 1, (pExt - pName) \ 2)
    
End Function

' // OpenFileName dialog
Public Function GetOpenFile( _
                ByVal hWnd As Long, _
                ByRef sTitle As String, _
                ByRef sFilter As String, _
                Optional ByRef sDefFileNameDir As String) As String
    Dim tOFN            As OPENFILENAME
    Dim strInputFile    As String
    Dim sDefDir         As String
    
    With tOFN
    
        .nMaxFile = 260
        strInputFile = String$(.nMaxFile, vbNullChar)
        
        If Len(sDefFileNameDir) Then
        
            memcpy ByVal StrPtr(strInputFile), ByVal StrPtr(sDefFileNameDir), LenB(sDefFileNameDir)
            sDefDir = sDefFileNameDir
            PathRemoveFileSpec sDefDir
            .lpstrInitialDir = StrPtr(sDefDir)
            
        End If
        
        .hwndOwner = hWnd
        .lpstrTitle = StrPtr(sTitle)
        .lpstrFile = StrPtr(strInputFile)
        .lStructSize = Len(tOFN)
        .lpstrFilter = StrPtr(sFilter)
        
        If GetOpenFileName(tOFN) = 0 Then Exit Function
        
        GetOpenFile = Left$(strInputFile, InStr(1, strInputFile, vbNullChar) - 1)
        
    End With

End Function

' // Show save file name dialog
Public Function GetSaveFile( _
                ByVal hWnd As Long, _
                ByRef sTitle As String, _
                ByRef sFilter As String, _
                ByRef sDefExtension As String, _
                Optional ByRef sDefFileName As String) As String
    Dim tOFN            As OPENFILENAME
    Dim strOutputFile   As String
    Dim sDefDir         As String
    
    With tOFN
    
        .nMaxFile = 260
        strOutputFile = String$(.nMaxFile, vbNullChar)
        
        If Len(sDefFileName) Then
            
            memcpy ByVal StrPtr(strOutputFile), ByVal StrPtr(sDefFileName), LenB(sDefFileName)
            sDefDir = sDefFileName
            PathRemoveFileSpec sDefDir
            .lpstrInitialDir = StrPtr(sDefDir)
            
        End If
        
        .hwndOwner = hWnd
        .lpstrTitle = StrPtr(sTitle)
        .lpstrFile = StrPtr(strOutputFile)
        .lStructSize = Len(tOFN)
        .lpstrFilter = StrPtr(sFilter)
        .lpstrDefExt = StrPtr(sDefExtension)
        .nFilterIndex = 1
        .flags = OFN_EXPLORER Or _
                 OFN_ENABLESIZING Or OFN_NOREADONLYRETURN Or OFN_PATHMUSTEXIST Or _
                 OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
                 
        If GetSaveFileName(tOFN) = 0 Then Exit Function
        
        GetSaveFile = Left$(strOutputFile, InStr(1, strOutputFile, vbNullChar) - 1)

    End With

End Function



