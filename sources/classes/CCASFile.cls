VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCASFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CCASFile.cls - this class manages file and converts between CAS/BAS/TXT
' // By The trick 2021
' //

Option Explicit
Option Base 0

Implements CBasicParser

Private Const MODULE_NAME = "CCASFile"

Private m_sFileName     As String
Private m_sSource       As String
Private m_sProgramName  As String
Private m_cParser       As CBasicParser
Private m_lOriginalHash As Long         ' // Original hash of data. Used to track changes

' // Current file name
Public Property Get FileName() As String
    FileName = m_sFileName
End Property
Public Property Let FileName( _
                    ByRef sValue As String)
    m_sFileName = sValue
End Property

' // Current source code
Public Property Get Source() As String
    Source = m_sSource
End Property
Public Property Let Source( _
                    ByRef sValue As String)
    m_sSource = FixUnicode(sValue)
End Property

' // Current program name
Public Property Get ProgramName() As String
    ProgramName = m_sProgramName
End Property
Public Property Let ProgramName( _
                    ByRef sValue As String)
    If Len(sValue) > 127 Then
        m_sProgramName = Left$(sValue, 127)
    Else
        m_sProgramName = sValue
    End If
End Property

' // Determines if there is unsaved changes
Public Property Get Changed() As Boolean
    Changed = CalcHash <> m_lOriginalHash
End Property

' // Load a CAS/BAS/TXT/KOI7 file
Public Sub Load( _
           ByRef sFileName As String)
    Const PROC_NAME = "Load", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bFileData()     As Byte
    Dim lFileSize       As Long
    Dim bBasicData()    As Byte
    Dim sProgramName    As String
    Dim cTxtFile        As CTextFile
    
    If Not LoadFileToArray(sFileName, bFileData, lFileSize) Then
        Err.Raise 75, FULL_PROC_NAME, "LoadFileToArray failed"
    End If
    
    Select Case LCase(GetFileExtension(sFileName))
    Case ".cas"
    
        If Not ConvertCASToBinBasic(bFileData, sProgramName, bBasicData) Then
            Err.Raise 5, FULL_PROC_NAME, "Invalid file format"
        End If
        
        m_sSource = m_cParser.ConvertBasicDataToString(bBasicData)
        m_sProgramName = sProgramName
        
    Case ".txt"
        
        Set cTxtFile = New CTextFile
        
        If lFileSize > 0 Then
            m_sSource = cTxtFile.LoadFromMemory(VarPtr(bFileData(0)), lFileSize)
        End If
        
        m_sProgramName = GetFileTitle(sFileName)
        
    Case ".bas"
        
        m_sSource = m_cParser.ConvertBasicDataToString(bFileData)
        m_sProgramName = GetFileTitle(sFileName)
        
    Case ".koi7"
    
        m_sSource = Vec6KOI7ToUnicode(bFileData)
        m_sProgramName = GetFileTitle(sFileName)
    
    End Select

    m_sFileName = sFileName
    m_lOriginalHash = CalcHash
    
End Sub

' // Save current data to CAS/BAS/TXT/KOI7
Public Sub Save( _
           ByRef sFileName As String)
    Const PROC_NAME = "Save", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bBasicData()    As Byte
    Dim bFileData()     As Byte
    Dim cTxtFile        As CTextFile
    
    Select Case LCase(GetFileExtension(sFileName))
    Case ".cas"
    
        m_cParser.ConvertSourceToBasicData m_sSource, bBasicData
        ConvertBinBasicToCAS bBasicData, m_sProgramName, bFileData
        
        If Not SaveArrayToFile(sFileName, bFileData) Then
            Err.Raise 75, FULL_PROC_NAME, "SaveArrayToFile failed"
        End If
        
    Case ".txt"
        
        Set cTxtFile = New CTextFile
        
        cTxtFile.Encoding = TE_UTF8 Or TE_HASBOM
        cTxtFile.Content = m_sSource
        
        If Not cTxtFile.SaveTextFile(sFileName) Then
            Err.Raise 75, FULL_PROC_NAME, "SaveTextFile failed"
        End If
        
    Case ".bas"
        
        m_cParser.ConvertSourceToBasicData m_sSource, bBasicData
        
        If Not SaveArrayToFile(sFileName, bBasicData) Then
            Err.Raise 75, FULL_PROC_NAME, "SaveArrayToFile failed"
        End If
        
    Case ".koi7"
        
        bFileData = UnicodeToVec6KOI7(m_sSource)
        
        If Not SaveArrayToFile(sFileName, bFileData) Then
            Err.Raise 75, FULL_PROC_NAME, "SaveArrayToFile failed"
        End If
        
    End Select

    m_sFileName = sFileName
    m_lOriginalHash = CalcHash
    
End Sub

' // Calculate hash on current data
Private Function CalcHash() As Long
    Const PROC_NAME = "CalcHash", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lHash   As Long
    
    If LenB(m_sSource) > 0 Then
        If HashData(ByVal StrPtr(m_sSource), LenB(m_sSource), lHash, Len(lHash)) < 0 Then
            Err.Raise 75, FULL_PROC_NAME, "HashData failed"
        End If
    Else
        lHash = 0
    End If
    
    CalcHash = lHash
    
    If LenB(m_sProgramName) > 0 Then
        If HashData(ByVal StrPtr(m_sProgramName), LenB(m_sProgramName), lHash, Len(lHash)) < 0 Then
            Err.Raise 75, FULL_PROC_NAME, "HashData failed"
        End If
    Else
        lHash = 0
    End If
    
    CalcHash = CalcHash Xor lHash
    
End Function

' // Convert raw basic toekns to CAS file
Private Sub ConvertBinBasicToCAS( _
            ByRef bData() As Byte, _
            ByRef sProgramName As String, _
            ByRef bOutData() As Byte)
    Const PROC_NAME = "ConvertBinBasicToCAS", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lSize           As Long
    Dim lBasSize        As Long
    Dim lIndex          As Long
    Dim lIndex2         As Long
    Dim bProgName()     As Byte
    Dim lProgNameSize   As Long
    Dim lCRC            As Long
    
    bProgName = UnicodeToVec6KOI7(sProgramName)
    lProgNameSize = SACount(ArrPtr(bProgName))
    lBasSize = SACount(ArrPtr(bData))
    lSize = lBasSize + 784 + lProgNameSize
    lCRC = CalcCRC(bData)
    
    ReDim bOutData(lSize - 1)
    
    For lIndex = 0 To 3
        bOutData(lIndex) = &HD3
    Next
    
    If Len(sProgramName) Then
        memcpy bOutData(4), bProgName(0), UBound(bProgName) + 1
        lIndex = lIndex + UBound(bProgName) + 1
    End If
    
    lIndex = lIndex + 3 ' // NULLs
    
    For lIndex2 = 0 To 767
        bOutData(lIndex) = &H55
        lIndex = lIndex + 1
    Next
    
    bOutData(lIndex) = &HE6
    lIndex = lIndex + 1
    
    For lIndex2 = 0 To 2
        bOutData(lIndex) = &HD3
        lIndex = lIndex + 1
    Next
    
    lIndex = lIndex + 1
    
    If lBasSize > 0 Then
        memcpy bOutData(lIndex), bData(0), lBasSize
    End If
    
    lIndex = lIndex + lBasSize + 2
    
    GetMem2 lCRC, bOutData(lIndex)
    
End Sub

' // Convert CAS file to raw basic tokens
Private Function ConvertCASToBinBasic( _
                 ByRef bData() As Byte, _
                 ByRef sProgramName As String, _
                 ByRef bOutData() As Byte) As Boolean
    Const PROC_NAME = "ConvertCASToBinBasic", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bOut()          As Byte
    Dim bProgramName()  As Byte
    Dim lIndex          As Long
    Dim lIndex2         As Long
    Dim lNameSize       As Long
    Dim lContentSize    As Long
    Dim lCRC            As Long
    Dim lSize           As Long
    
    lSize = SACount(ArrPtr(bData))
    
    If lSize < 784 Then
        Exit Function
    End If
    
    For lIndex = 0 To 3
        If bData(lIndex) <> &HD3 Then
            Exit Function
        End If
    Next
    
    Do While bData(lIndex) And lIndex < lSize - 1
        lIndex = lIndex + 1
    Loop
    
    If lIndex = lSize Then
        Exit Function
    End If
    
    lNameSize = lIndex - 4
    
    If lNameSize > 127 Then
        Exit Function
    End If
    
    lContentSize = lSize - lNameSize - 784
    
    If lNameSize > 0 Then
    
        ReDim bProgramName(lNameSize - 1)
        memcpy bProgramName(0), bData(4), lNameSize
    
    End If
    
    sProgramName = Vec6KOI7ToUnicode(bProgramName)
    
    For lIndex2 = 0 To 2
        
        If bData(lIndex) <> 0 Then Exit Function
        lIndex = lIndex + 1
        
    Next
    
    For lIndex2 = 0 To 767
        
        If bData(lIndex) <> &H55 Then Exit Function
        lIndex = lIndex + 1
        
    Next
    
    If bData(lIndex) <> &HE6 Then Exit Function
    lIndex = lIndex + 1
    
    For lIndex2 = 0 To 2
        
        If bData(lIndex) <> &HD3 Then Exit Function
        lIndex = lIndex + 1
        
    Next
    
    If bData(lIndex) <> 0 Then Exit Function
    lIndex = lIndex + 1
    
    If lContentSize > 0 Then
    
        ReDim bOut(lContentSize - 1)
    
        memcpy bOut(0), bData(lIndex), lContentSize
    
        lIndex = lIndex + lContentSize

    End If
    
    For lIndex2 = 0 To 1
        
        If bData(lIndex) <> 0 Then Exit Function
        lIndex = lIndex + 1
        
    Next
    
    ' // Check CRC
    lCRC = CalcCRC(bOut())
    
    If (lCRC And &HFF) <> bData(lIndex) Or _
        ((lCRC \ &H100) And &HFF) <> bData(lIndex + 1) Then Exit Function
    
    bOutData = bOut
    ConvertCASToBinBasic = True
    
End Function

' // Calculate CAS CRC
Private Function CalcCRC( _
                 ByRef bData() As Byte) As Long
    Dim lCRC    As Long
    Dim lIndex  As Long
    Dim lCount  As Long
    
    lCount = SACount(ArrPtr(bData))
    
    For lIndex = 0 To lCount - 1
        lCRC = (lCRC + bData(lIndex)) And &HFFFF&
    Next
    
    CalcCRC = lCRC
    
End Function

Private Function CBasicParser_ConvertBasicDataToString( _
                 ByRef bData() As Byte) As String
    CBasicParser_ConvertBasicDataToString = m_cParser.ConvertBasicDataToString(bData)
End Function

Private Sub CBasicParser_ConvertSourceToBasicData( _
            ByRef sSource As String, _
            ByRef bOut() As Byte)
    m_cParser.ConvertSourceToBasicData sSource, bOut
End Sub

Private Sub Class_Initialize()
    Set m_cParser = New CBasicParser
End Sub
