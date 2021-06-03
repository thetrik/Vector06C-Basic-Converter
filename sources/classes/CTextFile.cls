VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CTextFile.cls - this class can detect code-page of text and convert text to specified encoding
' // By The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME = "CTextFile"

Public Enum eTextEncoding
    TE_ANSI = 0
    TE_UNICODE = 1
    TE_UTF8 = 2
    TE_CPMASK = 3
    TE_BIGENDIAN = 4
    TE_HASBOM = 8
End Enum

Private m_sContent  As String
Private m_eEncoding As eTextEncoding
Private m_sFileName As String

Public Property Get Encoding() As eTextEncoding
    Encoding = m_eEncoding
End Property
Public Property Let Encoding( _
                    ByVal eValue As eTextEncoding)
    m_eEncoding = eValue
End Property

Public Property Get Content() As String
    Content = m_sContent
End Property
Public Property Let Content( _
                    ByRef sValue As String)
    m_sContent = sValue
End Property

Public Property Get FileName() As String
    FileName = m_sFileName
End Property
Public Property Let FileName( _
                    ByRef sValue As String)
    m_sFileName = sValue
End Property

' // Save current content to file using current encoding
Public Function SaveTextFile( _
                ByRef sFileName As String) As Boolean
    Const PROC_NAME = "SaveTextFile", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bData()     As Byte
    Dim hFile       As OLE_HANDLE
    Dim lSize       As Long
    Dim lIndex      As Long
    Dim lSymIdx     As Long
    Dim pString     As Long
    Dim lChar       As Long
    Dim lTotalSize  As Long
    Dim lCodePage   As Long
    
    On Error GoTo CleanUp
    
    hFile = CreateFile(sFileName, GENERIC_WRITE Or GENERIC_READ, 0, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Err.Raise 7, FULL_PROC_NAME, "CreateFile failed"
    End If
    
    Select Case (m_eEncoding And TE_CPMASK)
    Case TE_ANSI, TE_UTF8
        
        If (m_eEncoding And TE_CPMASK) = TE_ANSI Then
            lCodePage = CP_ACP
        Else
        
            lCodePage = CP_UTF8
                
            If m_eEncoding And TE_HASBOM Then
                lTotalSize = 3: lIndex = 3
            End If
        
        End If

        If Len(m_sContent) Then
        
            lSize = WideCharToMultiByte(lCodePage, 0, ByVal StrPtr(m_sContent), Len(m_sContent), ByVal 0&, 0, ByVal 0&, 0)
            If lSize = 0 Then
                Err.Raise 7, FULL_PROC_NAME, "WideCharToMultiByte failed"
            End If
            
            lTotalSize = lTotalSize + lSize
            
            ReDim bData(lTotalSize - 1)
            
            If WideCharToMultiByte(lCodePage, 0, ByVal StrPtr(m_sContent), Len(m_sContent), bData(lIndex), lSize, ByVal 0&, 0) = 0 Then
                Err.Raise 7, FULL_PROC_NAME, "WideCharToMultiByte failed"
            End If

        End If
        
        If (m_eEncoding And TE_HASBOM) And ((m_eEncoding And TE_CPMASK) = TE_UTF8) Then
            bData(0) = &HEF
            bData(1) = &HBB
            bData(2) = &HBF
        End If

    Case TE_UNICODE
        
        lSize = LenB(m_sContent)
    
        If m_eEncoding And TE_HASBOM Then
            lTotalSize = lSize + 2
        Else
            lTotalSize = lSize
        End If
        
        If lTotalSize > 0 Then
        
            ReDim bData(lTotalSize - 1)
            
            If m_eEncoding And TE_HASBOM Then
                If m_eEncoding And TE_BIGENDIAN Then
                    GetMem2 &HFFFE&, bData(0)
                Else
                    GetMem2 &HFEFF&, bData(0)
                End If
                
                lIndex = lIndex + 2
                
            End If
            
            If m_eEncoding And TE_BIGENDIAN Then
                
                pString = StrPtr(m_sContent)
            
                For lSymIdx = 0 To Len(m_sContent) - 1
                    GetMem2 ByVal pString + lSymIdx * 2, lChar
                    lChar = (lChar \ &H100) Or ((lChar And &HFF) * &H100)
                    GetMem2 lChar, bData(lIndex + lSymIdx * 2)
                Next
                
            Else
                memcpy bData(lIndex), ByVal StrPtr(m_sContent), LenB(m_sContent)
            End If
            
        End If
        
    End Select
       
    If lTotalSize > 0 Then
    
        If WriteFile(hFile, bData(0), lTotalSize, lSize, ByVal 0&) = 0 Then
            Err.Raise 7, FULL_PROC_NAME, "WriteFile failed"
        End If
        
        If lSize = lTotalSize Then
            SaveTextFile = True
        End If
        
    Else
        SaveTextFile = True
    End If
    
    m_sFileName = sFileName
    
CleanUp:
    
    CloseHandle hFile
    
    If Err.Number Then
        ThrowCurrentErrorUp FULL_PROC_NAME
    End If
    
End Function

' //
' // Load file and convert it to UTF-16
' //
Public Sub LoadTextFile( _
           ByRef sFileName As String)
    Const PROC_NAME = "LoadTextFile", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim hFile   As OLE_HANDLE
    Dim hMap    As OLE_HANDLE
    Dim pData   As Long
    Dim liSize  As LARGE_INTEGER
    
    hFile = CreateFile(sFileName, GENERIC_READ, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Err.Raise 7, FULL_PROC_NAME, "CreateFile failed"
    End If
    
    If GetFileSizeEx(hFile, liSize) = 0 Then
        CloseHandle hFile
        Err.Raise 7, FULL_PROC_NAME, "GetFileSizeEx failed"
    End If
    
    If liSize.HighPart <> 0 Or liSize.LowPart < 0 Or liSize.LowPart > 10000000 Then
        CloseHandle hFile
        Err.Raise 7, FULL_PROC_NAME, "File is too big"
    End If
    
    hMap = CreateFileMapping(hFile, ByVal 0&, PAGE_READONLY, 0, 0, vbNullString)
    CloseHandle hFile
    If hMap = 0 Then
        Err.Raise 7, FULL_PROC_NAME, "CreateFileMapping failed"
    End If
    
    pData = MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, 0)
    CloseHandle hMap
    If pData = 0 Then
        Err.Raise 7, FULL_PROC_NAME, "MapViewOfFile failed"
    End If
    
    On Error GoTo CleanUp
    
    LoadFromMemory pData, liSize.LowPart
    
    m_sFileName = sFileName
    
CleanUp:
    
    UnmapViewOfFile ByVal pData
       
    If Err.Number Then
        ThrowCurrentErrorUp FULL_PROC_NAME
    End If
       
End Sub

Public Function LoadFromMemory( _
                ByVal pData As Long, _
                ByVal lSize As Long) As String
    Const PROC_NAME = "LoadFromMemory", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim eEncoding   As eTextEncoding
    Dim sRet        As String
    Dim lRetSize    As Long
    Dim lIndex      As Long
    Dim lCodePage   As Long
    Dim lSwap       As Long
    Dim pRet        As Long
    
    If lSize = 0 Then Exit Function
    
    eEncoding = DetectEncoding(pData, lSize)

    If (eEncoding And TE_CPMASK) = TE_UNICODE Then
        
        If eEncoding And TE_HASBOM Then
            pData = pData + 2
            lSize = lSize - 2
        End If
        
        sRet = Space$((lSize + 1) \ 2)
        
        If eEncoding And TE_BIGENDIAN Then
            
            pRet = StrPtr(sRet)
            
            For lIndex = 0 To Len(sRet) - 1
                
                GetMem2 ByVal pData + lIndex * 2, lSwap
                lSwap = ((lSwap And &HFF) * &H100) Or (lSwap \ &H100)
                GetMem2 lSwap, ByVal pRet + lIndex * 2
                
            Next
            
        Else
            memcpy ByVal StrPtr(sRet), ByVal pData, lSize
        End If
        
    Else
        
        If (eEncoding And TE_CPMASK) = TE_ANSI Then
            lCodePage = CP_ACP
        ElseIf (eEncoding And TE_CPMASK) = TE_UTF8 Then
            
            lCodePage = CP_UTF8
            
            If eEncoding And TE_HASBOM Then
                pData = pData + 3
                lSize = lSize - 3
            End If
            
        End If
        
        lRetSize = MultiByteToWideChar(lCodePage, 0, ByVal pData, lSize, ByVal 0&, 0)
        If lRetSize = 0 Then
            Err.Raise 7, FULL_PROC_NAME, "MultiByteToWideChar failed"
        End If
        
        sRet = Space$(lRetSize)
        
        If MultiByteToWideChar(lCodePage, 0, ByVal pData, lSize, ByVal StrPtr(sRet), lRetSize) = 0 Then
            Err.Raise 7, FULL_PROC_NAME, "MultiByteToWideChar failed"
        End If
        
    End If
    
    m_sContent = sRet
    LoadFromMemory = sRet
    m_eEncoding = Encoding
    
End Function

Private Function DetectEncoding( _
                 ByVal pData As Long, _
                 ByVal lSize As Long) As eTextEncoding
    Dim lBOM    As Long
    Dim bBE     As Boolean
    
    If lSize < 2 Then
        DetectEncoding = TE_ANSI
        Exit Function
    End If
    
    GetMem2 ByVal pData, lBOM
    
    If lBOM = &HFEFF& Then
        ' // UTF-16 LE
        DetectEncoding = TE_UNICODE Or TE_HASBOM
    ElseIf lBOM = &HFFFE& Then
        ' // UTF-16 BE
        DetectEncoding = TE_UNICODE Or TE_BIGENDIAN Or TE_HASBOM
    ElseIf lSize > 2 Then
        If lBOM = &HBBEF& Then
            
            GetMem1 ByVal pData + 2, lBOM
            
            If (lBOM And &HFF) = &HBF Then
                ' // UTF-8
                DetectEncoding = TE_UTF8 Or TE_HASBOM
            End If
            
        Else
            If IsInputTextUnicode(pData, lSize, bBE) Then
                ' // UTF-16
                If bBE Then
                    DetectEncoding = TE_UNICODE Or TE_BIGENDIAN
                Else
                    DetectEncoding = TE_UNICODE
                End If
            ElseIf IsInputTextUTF8(pData, lSize) Then
                ' // UTF-8
                DetectEncoding = TE_UTF8
            Else
                ' // ANSI
                DetectEncoding = TE_ANSI
            End If
        End If
    Else
        DetectEncoding = TE_ANSI
    End If
                     
End Function

Private Function IsInputTextUTF8( _
                 ByVal pData As Long, _
                 ByVal lSize As Long) As Boolean
    Dim bChar   As Byte
    Dim lIndex  As Long
    Dim bNoHigh As Boolean
    Dim lCount  As Long
    
    If lSize <= 0 Then Exit Function
    
    bNoHigh = True
    
    For lIndex = 0 To lSize - 1
    
        GetMem1 ByVal pData + lIndex, bChar
        
        If (bChar And &H80) <> 0 Then
            bNoHigh = False
        End If
            
        If lCount Then
            
            If (bChar And &HC0) <> &H80 Then
                Exit Function
            End If
            
            lCount = lCount - 1
            
        ElseIf bChar >= &H80 Then
        
            Do
                
                bChar = (CLng(bChar) * 2) And &HFF
                lCount = lCount + 1
                
            Loop While bChar And &H80
            
            lCount = lCount - 1
            
            If lCount = 0 Then
                Exit Function
            End If
            
        End If
        
    Next
    
    If CBool(lCount) Or bNoHigh Then
        Exit Function
    Else
        IsInputTextUTF8 = True
    End If
    
End Function

Private Function IsInputTextUnicode( _
                 ByVal pData As Long, _
                 ByVal lSize As Long, _
                 ByRef bIsBigEndian As Boolean) As Boolean
    Dim lFlags  As Long
    
    lFlags = -1
    
    If IsTextUnicode(ByVal pData, lSize, lFlags) Then
        If lSize < 100 And lFlags = IS_TEXT_UNICODE_STATISTICS Then
            IsInputTextUnicode = False
        Else
        
            If lFlags = IS_TEXT_UNICODE_REVERSE_STATISTICS Then
                bIsBigEndian = True
            Else
                bIsBigEndian = False
            End If
            
            IsInputTextUnicode = True
            
        End If
    End If
    
End Function



