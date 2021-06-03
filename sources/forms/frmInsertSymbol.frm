VERSION 5.00
Begin VB.Form frmInsertSymbol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert symbol"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4230
   Icon            =   "frmInsertSymbol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInsertSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmInsertSymbol.frm - form with ability to add special basic symbols (0 - 31, 127 range)
' // By The trick 2021
' //

Option Explicit

Private Const MODULE_NAME = "frmInsertSymbol"

' // Button state
Private Enum eButtonState
    BS_PRESSED = 1  ' // Button was pressed
    BS_HOT = 2      ' // The mouse over button
    BS_DISABLED = 3 ' // Button was disabled
End Enum

' // It doesn't use controls at all. All the buttons are painted by code

Private m_lButtonWidth  As Long     ' // Buttons metrics
Private m_lButtonHeight As Long
Private m_lSpaceSize    As Long     ' // Space between form and buttons
Private m_lActiveButton As Long     ' // Current active button index. If no active then -1
Private m_cVectorFont   As StdFont
Private m_cOriginFont   As StdFont
Private m_tButtonsArea  As RECT     ' // Buttons area (to test mouse)

' // We don't need double-click
Private Sub Form_DblClick()
    Const PROC_NAME = "Form_DblClick", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo error_handler
    
    If m_lActiveButton >= 0 Then
        ButtonPressed m_lActiveButton
    End If
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_Load()
    Const PROC_NAME = "Form_Load", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo error_handler
    
    m_lActiveButton = -1
    
    Set m_cOriginFont = Me.Font
    Set m_cVectorFont = New StdFont
    
    ' // Big font
    m_cVectorFont.Name = "Vector06C"
    m_cVectorFont.Size = 12
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_MouseDown( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef x As Single, _
            ByRef y As Single)
    Const PROC_NAME = "Form_MouseDown", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lIndex  As Long
    
    On Error GoTo error_handler
    
    If PtInRect(m_tButtonsArea, x, y) Then
            
        lIndex = ButtonIndexFromCoord(x, y)
        
        If m_lActiveButton >= 0 Then
            DrawButton m_lActiveButton, 0
        End If
        
        ' // Skip CL and LF unprintable
        If lIndex = 10 Or lIndex = 13 Then
            m_lActiveButton = -1
            Exit Sub
        End If
        
        m_lActiveButton = lIndex
        
        DrawButton m_lActiveButton, BS_PRESSED
        
        ButtonPressed lIndex
            
    End If
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_MouseMove( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef x As Single, _
            ByRef y As Single)
    Const PROC_NAME = "Form_MouseMove", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lIndex  As Long
    
    On Error GoTo error_handler
    
    If PtInRect(m_tButtonsArea, x, y) Then

        lIndex = ButtonIndexFromCoord(x, y)
        
        If lIndex = m_lActiveButton Then
            Exit Sub
        End If
        
        If m_lActiveButton >= 0 Then
            DrawButton m_lActiveButton, 0
        End If
        
        If lIndex = 10 Or lIndex = 13 Then
            m_lActiveButton = -1
            Exit Sub
        End If
        
        m_lActiveButton = lIndex
        
        DrawButton m_lActiveButton, BS_HOT
        
    ElseIf m_lActiveButton >= 0 Then
    
        DrawButton m_lActiveButton, 0
        m_lActiveButton = -1
        
    End If
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_MouseUp( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef x As Single, _
            ByRef y As Single)
    Const PROC_NAME = "Form_MouseUp", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    On Error GoTo error_handler
    
    If m_lActiveButton >= 0 Then
    
        DrawButton m_lActiveButton, 0
        m_lActiveButton = -1
        
    End If
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_Paint()
    Const PROC_NAME = "Form_Paint", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    Dim lIndex  As Long
    Dim eState  As eButtonState
    Dim lX      As Long
    Dim lY      As Long
    
    On Error GoTo error_handler
    
    For lIndex = 0 To 31
    
        If lIndex = 10 Or lIndex = 13 Then
            eState = BS_DISABLED
        Else
            eState = 0
        End If
        
        DrawButton lIndex, eState
            
    Next
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

' // The event when a button was pressed
Private Sub ButtonPressed( _
            ByVal lIndex As Long)
    Const PROC_NAME = "ButtonPressed", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bText(0) As Byte
    
    On Error GoTo error_handler
    
    If lIndex = 0 Then
        lIndex = 127    ' // Button with 0 index = 127 char
    End If
    
    bText(0) = lIndex
    
    frmMain.ctlTextBox.SelText = Vec6KOI7ToUnicode(bText)
 
    Exit Sub
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
          
End Sub

' // Calculate index from ccords
Private Function ButtonIndexFromCoord( _
                 ByVal lX As Long, _
                 ByVal lY As Long) As Long
    Dim lCol    As Long
    Dim lRow    As Long
    
    lX = lX - m_lSpaceSize
    lY = lY - m_lSpaceSize
        
    lRow = lX \ m_lButtonWidth
    lCol = lY \ m_lButtonHeight
        
    ButtonIndexFromCoord = lCol * 8 + lRow
        
End Function

Private Sub DrawButton( _
            ByVal lIndex As Long, _
            ByVal eState As eButtonState)
    Const PROC_NAME = "DrawButton", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lX      As Long
    Dim lY      As Long
    Dim lRow    As Long
    Dim lCol    As Long
    Dim tRC     As RECT
    Dim sText   As String
    Dim bKOI(0) As Byte
    
    On Error GoTo error_handler
    
    lRow = lIndex Mod 8
    lCol = lIndex \ 8

    lX = m_lSpaceSize + lRow * m_lButtonWidth
    lY = m_lSpaceSize + lCol * m_lButtonHeight
    
    Me.Line (lX, lY)-Step(m_lButtonWidth, m_lButtonHeight), Me.BackColor, BF
    
    Select Case eState
    Case BS_PRESSED
        Me.Line (lX + 1, lY + 1)-Step(m_lButtonWidth - 2, m_lButtonHeight - 2), vbHighlight, BF
    Case BS_HOT
        Me.Line (lX + 1, lY + 1)-Step(m_lButtonWidth - 2, m_lButtonHeight - 2), vb3DHighlight, B
    Case BS_DISABLED
        Me.Line (lX + 1, lY + 1)-Step(m_lButtonWidth - 2, m_lButtonHeight - 2), vbInactiveBorder, B
    Case Else
        Me.Line (lX + 1, lY + 1)-Step(m_lButtonWidth - 2, m_lButtonHeight - 2), vbActiveBorder, B
    End Select
    
    SetRect tRC, lX, lY, lX + m_lButtonWidth, lY + m_lButtonHeight
    
    If lIndex = 10 Or lIndex = 13 Then
        sText = " "
    Else
    
        If lIndex = 0 Then
            lIndex = 127
        End If
        
        bKOI(0) = lIndex
        sText = Vec6KOI7ToUnicode(bKOI)
        
    End If
    
    Set Me.Font = m_cVectorFont
    
    ' // Draw symbol
    DrawText Me.hDC, sText, -1, tRC, DT_CALCRECT

    OffsetRect tRC, (m_lButtonWidth - (tRC.Right - tRC.Left)) \ 2, (m_lButtonHeight - (tRC.Bottom - tRC.Top)) \ 2
    
    DrawText Me.hDC, sText, -1, tRC, 0
    
    Set Me.Font = m_cOriginFont
    
    ' // Draw hex code
    OffsetRect tRC, 0, tRC.Bottom - tRC.Top
    
    sText = Hex$(lIndex)
    
    If Len(sText) = 1 Then
        sText = "0" & sText
    End If
    
    DrawText Me.hDC, sText, -1, tRC, DT_CENTER
 
    Exit Sub
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
       
End Sub

Private Sub Form_Resize()
    Const PROC_NAME = "Form_Resize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lTextWidth  As Long
    Dim lTextHeight As Long
    
    On Error GoTo error_handler
    
    lTextWidth = Me.TextWidth("0")
    lTextHeight = Me.TextHeight("0")
    
    If lTextWidth > lTextHeight Then
        m_lSpaceSize = lTextWidth
    Else
        m_lSpaceSize = lTextHeight
    End If
    
    m_lSpaceSize = m_lSpaceSize
    
    m_lButtonWidth = (Me.ScaleWidth - m_lSpaceSize * 2) \ 8
    m_lButtonHeight = (Me.ScaleHeight - m_lSpaceSize * 2) \ 4
    
    SetRect m_tButtonsArea, m_lSpaceSize, m_lSpaceSize, _
            m_lButtonWidth * 8 + m_lSpaceSize, m_lButtonHeight * 4 + m_lSpaceSize
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub
