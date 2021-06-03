VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParserTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CParserTree.cls - this class can convert a text code representaion to tokens
' // By The trick 2021
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME = "CParserTree"

Public Enum eTokenType
    TT_NEWLINE = 0
    TT_REM = 142
    TT_DATA = 131
End Enum

' // This is parser tree
' //
' // [R]->[E]->[S]->[T]->[O]->[R]->[E]
' //  |    |
' //  |    +->[T]->[U]->[R]->[N]
' //  +->[U]->[N]
' //

' // The first 255 items contains ASCII symbols mapping
Private Type tTreeItem
    bSymbol     As Byte
    bToken      As Byte
    lParent     As Long
    lCount      As Long
    lChildren() As Long ' // Child indices
End Type

Private m_tItems()  As tTreeItem    ' // It has 255 items in 1st level
Private m_lPointer  As Long         ' // Index of item which will be added
Private m_lCount    As Long         ' // Count of items

Private m_lCurIndex As Long     ' // Current tree item index. -1 if initialized
Private m_bInQuote  As Boolean  ' // The parser is in quote (process symbols without keywords)
Private m_bRemData  As Boolean  ' // The parser si in REM/DATA block

' // Reset state
Public Sub Reset()
    
    m_lCurIndex = -1
    m_bInQuote = False
    m_bRemData = False
    
End Sub

' // Check if parsing process is incomplete
Public Property Get Complete() As Boolean
    Complete = m_lCurIndex = -1
End Property

' // Put a symbol to parser
Public Function PutSymbol( _
                ByVal bSymbol As Byte) As Byte()
    Dim lIndex  As Long
    Dim bRet()  As Byte
    Dim lSize   As Long

    ' // Check keywords
    If m_lCurIndex = -1 Then
        
        ' // 1st level
        m_lCurIndex = bSymbol
        
        If bSymbol = 34 And Not m_bRemData Then
            m_bInQuote = Not m_bInQuote
        End If
            
        Exit Function

    Else
    
        For lIndex = 0 To m_tItems(m_lCurIndex).lCount - 1
            With m_tItems(m_tItems(m_lCurIndex).lChildren(lIndex))
                If .bSymbol = bSymbol Then
                    If .lCount = 0 Then
                        
                        If .bToken = TT_NEWLINE Then
                            m_bRemData = False
                            m_bInQuote = False
                        Else
                            If Not m_bRemData Then
                                If .bToken = TT_DATA Or .bToken = TT_REM Then
                                    m_bRemData = True
                                ElseIf .bSymbol = 34 Or m_bInQuote Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                        
                        m_lCurIndex = -1
                        
                        ReDim bRet(0)
                        bRet(0) = .bToken
                        PutSymbol = bRet
                        
                        Exit Function
                        
                    Else

                        m_lCurIndex = m_tItems(m_lCurIndex).lChildren(lIndex)
                        Exit Function
                        
                    End If
                End If
            End With
        Next
        
        With m_tItems(m_lCurIndex)
            
            If .bToken Then
                
                ' // Existing keyword was broken
                ReDim bRet(0)
                
                If m_bInQuote Or m_bRemData Then
                    bRet(0) = .bSymbol
                Else
                    bRet(0) = .bToken
                End If
                
                If .bToken = TT_DATA Or .bToken = TT_REM Then
                    m_bRemData = True
                End If

            ElseIf .lCount Then
                
                ' // Part of keyword, translate to symbols
                lSize = NumberOfParents(m_lCurIndex)
                
                ReDim bRet(lSize - 1)
                
                Do While m_lCurIndex > 0
                
                    lSize = lSize - 1
                    bRet(lSize) = m_tItems(m_lCurIndex).bSymbol
                    m_lCurIndex = m_tItems(m_lCurIndex).lParent
                    
                Loop
            
            End If
            
            If bSymbol = 34 And Not m_bRemData Then
                m_bInQuote = Not m_bInQuote
            End If
            
            PutSymbol = bRet
            m_lCurIndex = bSymbol
                
        End With
        
    End If
    
End Function

' // Add a token
Private Sub Add( _
            ByRef sValue As String, _
            ByVal bToken As Byte)
    Const PROC_NAME = "Add", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bValue()    As Byte
    Dim lIndex      As Long
    Dim lNewIndex   As Long
    
    If Len(sValue) = 0 Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    bValue = UnicodeToVec6KOI7(sValue)
    
    lNewIndex = -1
    
    For lIndex = 0 To UBound(bValue)
        ' // Populate tree
        lNewIndex = EnsureInTree(lNewIndex, bValue(lIndex))
    Next
    
    m_tItems(lNewIndex).bToken = bToken
    
End Sub

Private Property Get Count() As Long
    Count = m_lCount
End Property

' // Returns number of parents of item
Private Function NumberOfParents( _
                 ByVal lIndex As Long) As Long

    Do While lIndex > 0
        lIndex = m_tItems(lIndex).lParent
        NumberOfParents = NumberOfParents + 1
    Loop
    
End Function

' // Return index of child item or add new if there is no item
Private Function EnsureInTree( _
                 ByVal lIndex As Long, _
                 ByVal bValue As Byte) As Long
    Dim lNewIndex   As Long
    Dim lChildIndex As Long
    Dim lIndex2     As Long
    
    ' // To 1st level
    If lIndex = -1 Then
    
        If m_tItems(bValue).bSymbol = 0 Then
            m_tItems(bValue).bSymbol = bValue
            m_lCount = m_lCount + 1
        End If
        
        EnsureInTree = bValue
        Exit Function
        
    End If
    
    lChildIndex = m_tItems(lIndex).lCount
    
    ' // Check if already exists
    For lIndex2 = 0 To lChildIndex - 1
        If m_tItems(m_tItems(lIndex).lChildren(lIndex2)).bSymbol = bValue Then
            EnsureInTree = m_tItems(lIndex).lChildren(lIndex2)
            Exit Function
        End If
    Next
    
    ' // Add new
    If lChildIndex = 0 Then
        ReDim m_tItems(lIndex).lChildren(9)
    Else
        If lChildIndex > UBound(m_tItems(lIndex).lChildren) Then
            ReDim Preserve m_tItems(lIndex).lChildren(lChildIndex + 10)
        End If
    End If
    
    lNewIndex = m_lPointer
    
    If lNewIndex > UBound(m_tItems) Then
        ReDim Preserve m_tItems(lNewIndex + 256)
    End If
    
    m_tItems(lNewIndex).bSymbol = bValue
    m_tItems(lNewIndex).lParent = lIndex
    m_tItems(lIndex).lChildren(lChildIndex) = lNewIndex
    m_tItems(lIndex).lCount = m_tItems(lIndex).lCount + 1
    
    m_lPointer = m_lPointer + 1
    m_lCount = m_lCount + 1
    
    EnsureInTree = lNewIndex
    
End Function

Private Sub Class_Initialize()
    Dim lIndex  As Long
    
    ReDim m_tItems(255)
    
    For lIndex = 0 To 255
        m_tItems(lIndex).lParent = -1
    Next
    
    m_lPointer = 256
    
    For lIndex = 1 To 31
        Add Chr$(lIndex), lIndex
    Next
    
    Add vbNewLine, TT_NEWLINE
    
    For lIndex = 32 To 228
        Add Vec6BasicKeyword(lIndex), lIndex
    Next
    
End Sub
