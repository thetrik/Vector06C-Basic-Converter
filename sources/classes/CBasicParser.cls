VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBasicParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CBasicParser.cls - this class allow to convert between Vector-06C BASIC source and tokens
' // By The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME = "CBasicParser"

Private Type tSourceLine
    lNextAddr   As Long
    lLineNumber As Long
    lTokens     As Long
    bTokens()   As Byte
End Type

Private m_cParser   As CParserTree

' // Convert tokenized basic data to source code
Public Function ConvertBasicDataToString( _
                ByRef bData() As Byte) As String
    Const PROC_NAME = "ConvertBasicDataToString", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lIndex      As Long
    Dim lNextAddr   As Long
    Dim lMode       As Long
    Dim lLine       As Long
    Dim sRet        As String
    Dim lSize       As Long
    
    lSize = SACount(ArrPtr(bData))
    
    lMode = 0   ' // Line number
    
    Do
        
        Select Case lMode
        Case 0  ' // Next line address low
            lNextAddr = bData(lIndex)
            lMode = 1
        Case 1  ' // Next line address high
        
            lNextAddr = (bData(lIndex) * &H100&) Or lNextAddr
            
            If lNextAddr = 0 Then Exit Do
            
            lMode = 2
        Case 2  ' // Line number parsing low
            lLine = bData(lIndex)
            lMode = 3
        Case 3  ' // Line number parsing high
        
            lLine = (bData(lIndex) * &H100&) Or lLine
            lMode = 4
            sRet = sRet & CStr(lLine) & " "
            
        Case Else   ' // Tokens
            
            Select Case bData(lIndex)
            Case 0
                sRet = sRet & vbNewLine
                lMode = 0
            Case Is <= 228
                sRet = sRet & Vec6BasicKeyword(bData(lIndex))
            Case Else
                Stop    ' // Unexpected
            End Select
            
        End Select

        lIndex = lIndex + 1
        
    Loop While lIndex < lSize
    
    ConvertBasicDataToString = sRet

End Function

' // Convert source code to tokenized basic data
Public Sub ConvertSourceToBasicData( _
           ByRef sSource As String, _
           ByRef bOut() As Byte)
    Const PROC_NAME = "ConvertSourceToBasicData", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bSource()   As Byte
    Dim tLines()    As tSourceLine
    Dim lLines      As Long
    Dim lChrIndex   As Long
    Dim lOutSize    As Long
    Dim lIndex      As Long
    Dim lLineIndex  As Long
    Dim lAddress    As Long
    
    If Len(sSource) = 0 Then
        Erase bOut
        Exit Sub
    End If
    
    bSource = UnicodeToVec6KOI7(sSource)
    lAddress = &H4301   ' // Start basic program address
    
    Do While lChrIndex <= UBound(bSource)
        
        If lLines Then
            If lLines > UBound(tLines) Then
                ReDim Preserve tLines(lLines * 2 - 1)
            End If
        Else
            ReDim tLines(49)
        End If
        
        lChrIndex = ParseLine(bSource(), lChrIndex, tLines(lLines))
        
        lAddress = lAddress + tLines(lLines).lTokens + 4
        
        tLines(lLines).lNextAddr = lAddress
        
        lOutSize = lOutSize + tLines(lLines).lTokens
        lLines = lLines + 1
        
    Loop
    
    lOutSize = lOutSize + lLines * 4
    
    If lOutSize > 0 Then
        ReDim bOut(lOutSize - 1)
    Else
        Erase bOut
    End If
    
    For lLineIndex = 0 To lLines - 1

        With tLines(lLineIndex)
            
            GetMem2 .lNextAddr, bOut(lIndex)
            lIndex = lIndex + 2
            
            GetMem2 .lLineNumber, bOut(lIndex)
            lIndex = lIndex + 2
        
            memcpy bOut(lIndex), .bTokens(0), .lTokens
            lIndex = lIndex + .lTokens
            
        End With
        
    Next
    
End Sub

' // Parse a text line
Private Function ParseLine( _
                 ByRef bData() As Byte, _
                 ByVal lIndex As Long, _
                 ByRef tOut As tSourceLine) As Long
    Const PROC_NAME = "ParseLine", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim eTokens()   As Byte
    Dim lTokenIndex As Long
    Dim lCount      As Long
    Dim bDone       As Boolean
    Dim bChar       As Byte
    
    tOut.lLineNumber = 0
    tOut.lTokens = 0
    
    Do While lIndex <= UBound(bData)
        
        Select Case bData(lIndex)
        Case &H30 To &H39
            
            tOut.lLineNumber = tOut.lLineNumber * 10 + bData(lIndex) - &H30
            
            If tOut.lLineNumber > 65535 Then
                Err.Raise 5, FULL_PROC_NAME, "Error line number"
            End If
            
        Case Else
            
            If lIndex = 0 Then
                Err.Raise 5, FULL_PROC_NAME, "Line number not found"
            End If
            
            Exit Do
            
        End Select

        lIndex = lIndex + 1
        
    Loop
    
    m_cParser.Reset

    ' // Skip spaces
    Do While lIndex <= UBound(bData)
    
        If bData(lIndex) <> 32 Then
            Exit Do
        End If
        
        lIndex = lIndex + 1
        
    Loop
    
    'Debug.Assert tOut.lLineNumber <> 480
    
    Do Until bDone
        
        If lIndex > UBound(bData) Then
            If m_cParser.Complete Then
                Exit Do
            Else
                If lIndex = UBound(bData) + 1 Then
                    bChar = 13
                Else
                    bChar = 10
                End If
            End If
        Else
            bChar = bData(lIndex)
        End If
        
        eTokens = m_cParser.PutSymbol(bChar)
        
        lCount = SACount(ArrPtr(eTokens))
        
        If lCount > 0 Then

            If tOut.lTokens Then
                If tOut.lTokens + lCount > UBound(tOut.bTokens) Then
                    ReDim Preserve tOut.bTokens(tOut.lTokens * 2 + lCount - 1)
                End If
            Else
                ReDim tOut.bTokens(31)
            End If
            
            memcpy tOut.bTokens(tOut.lTokens), eTokens(0), lCount
            tOut.lTokens = tOut.lTokens + lCount
             
            If eTokens(0) = TT_NEWLINE Then
                lIndex = lIndex + 1
                Exit Do
            End If
            
        End If
        
        lIndex = lIndex + 1
        
    Loop

    ParseLine = lIndex
    
End Function

Private Sub Class_Initialize()
    Set m_cParser = New CParserTree
End Sub
