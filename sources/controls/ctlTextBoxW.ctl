VERSION 5.00
Begin VB.UserControl ctlTextBox 
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // ctlTextBox.cls - simple UNICODE textbox control
' // By The trick 2021
' //

Option Explicit
Option Base 0

Private WithEvents m_cFont As StdFont
Attribute m_cFont.VB_VarHelpID = -1

Private m_hWnd          As Long
Private m_hActualFont   As Long ' // Because StdFont uses antialiasing the control re-creates the font without antialiasing

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor( _
                    ByVal lValue As OLE_COLOR)
                    
    UserControl.BackColor = lValue
    InvalidateRect m_hWnd, ByVal 0&, 1
    PropertyChanged "BackColor"
    
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor( _
                    ByVal lValue As OLE_COLOR)
                    
    UserControl.ForeColor = lValue
    InvalidateRect m_hWnd, ByVal 0&, 1
    PropertyChanged "ForeColor"
    
End Property

Public Property Get Font() As StdFont
    Set Font = m_cFont
End Property

Public Property Set Font( _
                    ByVal cValue As StdFont)
    Set m_cFont = cValue
    UpdateFont
    PropertyChanged "Font"
End Property

Public Property Get Text() As String
    Dim lSize   As Long
    
    lSize = GetWindowTextLength(m_hWnd)
    
    If lSize > 0 Then
        Text = Space$(lSize)
        GetWindowText m_hWnd, Text, lSize + 1
    End If
    
End Property

Public Property Let Text( _
                    ByRef sValue As String)
    SetWindowText m_hWnd, sValue
End Property
       
Public Property Let SelStart( _
                    ByVal lValue As Long)
    Dim lCurStart   As Long
    Dim lCurEnd     As Long
    
    SendMessage m_hWnd, EM_GETSEL, VarPtr(lCurStart), lCurEnd
    SendMessage m_hWnd, EM_SETSEL, lValue, ByVal lCurEnd
    
End Property
      
Public Property Get SelStart() As Long
    SendMessage m_hWnd, EM_GETSEL, VarPtr(SelStart), ByVal 0&
End Property
      
Public Property Let SelLength( _
                    ByVal lValue As Long)
    Dim lCurStart   As Long
    
    SendMessage m_hWnd, EM_GETSEL, VarPtr(lCurStart), ByVal 0&
    SendMessage m_hWnd, EM_SETSEL, lCurStart, ByVal lCurStart + lValue
                
End Property

Public Property Get SelLength() As Long
    Dim lCurStart   As Long
    Dim lCurEnd     As Long
    
    SendMessage m_hWnd, EM_GETSEL, VarPtr(lCurStart), lCurEnd
    
    SelLength = lCurEnd - lCurStart
    
End Property
      
Public Property Let SelText( _
                    ByRef sValue As String)
    SendMessage m_hWnd, EM_REPLACESEL, 1, ByVal StrPtr(sValue)
End Property

Public Sub SelectAll()
    SendMessage m_hWnd, EM_SETSEL, 0, ByVal -1&
End Sub

Private Sub m_cFont_FontChanged( _
            ByVal PropertyName As String)
    UpdateFont
End Sub

Private Sub UpdateFont()
    Dim cFont       As IFont
    Dim tLogFont    As LOGFONT
    Dim hNewFont    As Long
    
    Set cFont = m_cFont
    
    If GetObjectAPI(cFont.hFont, LenB(tLogFont), tLogFont) Then
        
        tLogFont.lfQuality = NONANTIALIASED_QUALITY
        
        hNewFont = CreateFontIndirect(tLogFont)
        
        If hNewFont Then
            
            If m_hActualFont Then
                DeleteObject m_hActualFont
            End If
            
            m_hActualFont = hNewFont
            
        End If
        
    End If
    
    SendMessage m_hWnd, WM_SETFONT, m_hActualFont, 1
    
End Sub

Private Sub UserControl_Initialize()

    m_hWnd = CreateWindowEx(WS_EX_CLIENTEDGE, _
                            "Edit", vbNullString, WS_CHILD Or WS_VISIBLE Or ES_AUTOHSCROLL Or _
                            ES_AUTOVSCROLL Or ES_MULTILINE Or ES_WANTRETURN Or WS_VSCROLL, _
                            0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
    SendMessage m_hWnd, EM_LIMITTEXT, 0, ByVal 0&
    
End Sub

Private Sub UserControl_GotFocus()
    SetFocusAPI m_hWnd
End Sub

Private Sub UserControl_Resize()
    MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
End Sub

Private Sub UserControl_InitProperties()
    Set m_cFont = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties( _
            ByRef PropBag As PropertyBag)
            
    Set m_cFont = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    UpdateFont
    
End Sub

Private Sub UserControl_Terminate()
    If m_hActualFont Then
        DeleteObject m_hActualFont
    End If
End Sub

Private Sub UserControl_WriteProperties( _
            ByRef PropBag As PropertyBag)
            
    PropBag.WriteProperty "Font", m_cFont, Ambient.Font
    PropBag.WriteProperty "BackColor", UserControl.BackColor, vbWindowBackground
    PropBag.WriteProperty "ForeColor", UserControl.ForeColor, Ambient.ForeColor
    
End Sub
