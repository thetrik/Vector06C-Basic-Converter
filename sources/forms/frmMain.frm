VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   4485
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5505
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin Vector06CBasic.ctlTextBox ctlTextBox 
      Height          =   4335
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Vector06C"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save as..."
      End
      Begin VB.Menu mnuProgramName 
         Caption         =   "Program name..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuInsertSymbol 
         Caption         =   "&Insert symbol..."
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings..."
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' //
' // frmMain.frm - main window and GUI-logic
' // By The trick 2021
' //

Option Explicit

Private Const MODULE_NAME = "frmMain"

Implements ISubclass        ' // Form supports sublcassing

Private m_cFile As CCASFile ' // Current file
Private m_hFont As Long     ' // Handle of installed Vector06C font

' // Set current theme of textbox
' // The frmSettings uses this method to apply settings
' // This method should be public
Public Sub SetTextboxTheme( _
           ByVal eTheme As eTheme)
    Const PROC_NAME = "SetTextboxTheme", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lForeColor  As OLE_COLOR
    Dim lBackColor  As OLE_COLOR
    
    On Error GoTo error_handler
    
    Select Case eTheme
    Case T_WIN
        lForeColor = vbWindowText
        lBackColor = vbWindowBackground
    Case T_DARK
        lBackColor = &H580000
        lForeColor = &HD9D9&
    Case T_SOFT
        lBackColor = &HDBCDBF
        lForeColor = &H303030
    End Select
    
    ctlTextBox.BackColor = lBackColor
    ctlTextBox.ForeColor = lForeColor
    
    Exit Sub
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Sub

' // Set current size of textbox
' // The frmSettings uses this method to apply settings
' // This method should be public
Public Sub SetTextboxFontSize( _
           ByVal eSize As eFontSize)
    Const PROC_NAME = "SetTextboxFontSize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo error_handler
    
    Select Case eSize
    Case FS_SMALL
        ctlTextBox.Font.Size = 6    ' // The font developed with 6 size so all sizes should be multiple this value
    Case FS_LARGE
        ctlTextBox.Font.Size = 12
    End Select
    
    Exit Sub
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Sub

Private Sub Form_Initialize()
    Const PROC_NAME = "Form_Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim bFontData() As Byte
    
    On Error GoTo error_handler
    
    ' // Install private font (Vector06C) from resources
    bFontData = LoadResData(100, RT_FONT)
    
    m_hFont = AddFontMemResourceEx(bFontData(0), UBound(bFontData) + 1, 0, 0)
    
    If m_hFont = 0 Then
        Err.Raise 7, PROC_NAME, "AddFontMemResourceEx failed"
    End If
    
    Exit Sub
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Sub

Private Sub Form_Load()
    Const PROC_NAME = "Form_Initialize", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME

    On Error GoTo error_handler
    
    ' // Subclass me
'    If Not SubclassWindow(Me) Then
'        Err.Raise 7, FULL_PROC_NAME, "SubclassWindow failed"
'    End If
        
    Set m_cFile = New CCASFile
    
    ' // Update textbox settings from ini file
    SetTextboxTheme Settings(SET_THEME)
    SetTextboxFontSize Settings(SET_FONT_SIZE)
    
    ' // Update info about file
    UpdateInfo
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_QueryUnload( _
            ByRef Cancel As Integer, _
            ByRef UnloadMode As Integer)

    On Error GoTo error_handler
    
    ' // Try to close current file. It'll show dialog if file was changed
    If Not CloseCurrentFile() Then
        Cancel = True
    End If
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_Resize()

    On Error GoTo error_handler
    
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    ctlTextBox.Move ctlTextBox.Left, ctlTextBox.Top, Me.ScaleWidth - ctlTextBox.Left * 2, _
                                    Me.ScaleHeight - ctlTextBox.Top * 2
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_Terminate()

    On Error GoTo error_handler
    
    ' // Remove installed font if it was installed
    If m_hFont Then
        RemoveFontMemResourceEx m_hFont
    End If
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)

    On Error GoTo error_handler
    
    UnsubclassWindow Me
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Property Get ISubclass_hWnd() As Long
    Const PROC_NAME = "ISubclass_OnWindowProc", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo error_handler
    
    ISubclass_hWnd = Me.hWnd
    
    Exit Property
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Property

Private Function ISubclass_OnWindowProc( _
                 ByVal hWnd As Long, _
                 ByVal lMsg As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Long, _
                 ByRef bDefCall As Boolean) As Long
    Const PROC_NAME = "ISubclass_OnWindowProc", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo error_handler
    
    Dim tMinMaxInfo As MINMAXINFO
    
    Select Case lMsg
    ' // Process minimum/maximum size of window
    Case WM_GETMINMAXINFO
    
        memcpy tMinMaxInfo, ByVal lParam, Len(tMinMaxInfo)
        
        tMinMaxInfo.ptMaxTrackSize.x = Screen.Width / Screen.TwipsPerPixelX
        tMinMaxInfo.ptMaxTrackSize.y = Screen.Height / Screen.TwipsPerPixelY
        tMinMaxInfo.ptMinTrackSize.x = 320
        tMinMaxInfo.ptMinTrackSize.y = 240
        
        memcpy ByVal lParam, tMinMaxInfo, Len(tMinMaxInfo)
        
        bDefCall = False
        
    Case Else
        bDefCall = True
    End Select
    
    Exit Function
    
error_handler:
    
    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Function

Private Sub mnuAbout_Click()

    On Error GoTo error_handler
    
    frmAbout.Show vbModal, Me
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub mnuInsertSymbol_Click()
    
    On Error GoTo error_handler
    
    frmInsertSymbol.Show vbModeless, Me
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub mnuOpen_Click()
    Dim sFileName   As String
    
    On Error GoTo error_handler
    
    If Not CloseCurrentFile Then Exit Sub
    
    sFileName = GetOpenFile(Me.hWnd, "Open file", "All supported files" & vbNullChar & "*.cas;*.txt;*.bas;*.koi7" & vbNullChar)
    If Len(sFileName) = 0 Then Exit Sub
    
    m_cFile.Load sFileName
    
    ctlTextBox.Text = m_cFile.Source
    
    UpdateInfo
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

' // Save current file
' // Returns true if successful
Private Function SaveCurrentFile( _
                 ByVal bSaveAs As Boolean) As Boolean
    Const PROC_NAME = "SaveCurrentFile", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim sFileName   As String

    On Error GoTo err_handler
    
    m_cFile.Source = ctlTextBox.Text
    
    ' // Show dialog if this is new file or Save as... menu was selected
    If Len(m_cFile.FileName) And Not bSaveAs Then
        m_cFile.Save m_cFile.FileName
    Else
    
        sFileName = GetSaveFile(Me.hWnd, "Save file", "CAS files" & vbNullChar & "*.cas" & vbNullChar & _
                                "Text files" & vbNullChar & "*.txt" & vbNullChar & _
                                "BAS files" & vbNullChar & "*.bas" & vbNullChar & _
                                "KOI-7 N2 files" & vbNullChar & "*.koi7" & vbNullChar, _
                                "cas", GetFileTitle(m_cFile.FileName))
        If Len(sFileName) = 0 Then Exit Function
    
        m_cFile.Save sFileName
        
    End If
    
    UpdateInfo
    
    SaveCurrentFile = True
    
    Exit Function
    
err_handler:

    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Function

' // Show form caption with program name and file name
Private Sub UpdateInfo()
    Const PROC_NAME = "UpdateInfo", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim sFileName   As String
    
    On Error GoTo err_handler
    
    sFileName = m_cFile.FileName
    
    If Len(sFileName) Then
        sFileName = "(" & sFileName & ")"
    End If
    
    Me.Caption = "VECTOR-06C BASIC converter by The trick " & sFileName
    
    Exit Sub
    
err_handler:

    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Sub

' // Close current file
' // If file was changed since last saving it shows Save dialog
' // Returns true if file can be closed
Private Function CloseCurrentFile() As Boolean
    Const PROC_NAME = "CloseCurrentFile", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    On Error GoTo err_handler
    
    m_cFile.Source = ctlTextBox.Text
    
    If m_cFile.Changed Then
        Select Case MsgBox("The file was modified. Do you want to save changes?", vbYesNoCancel Or vbQuestion)
        Case vbYes
            CloseCurrentFile = SaveCurrentFile(False)
        Case vbNo
            CloseCurrentFile = True
        End Select
    Else
        CloseCurrentFile = True
    End If
    
    Exit Function
    
err_handler:

    ThrowCurrentErrorUp FULL_PROC_NAME
    
End Function

Private Sub mnuProgramName_Click()
    Dim sProgramName    As String
    
    On Error GoTo err_handler
    
    sProgramName = InputBox("Enter program name:", , m_cFile.ProgramName)
    If StrPtr(sProgramName) = 0 Then Exit Sub
    
    m_cFile.ProgramName = sProgramName
    
    Exit Sub
    
err_handler:

    ShowCurrentError
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()

    On Error GoTo err_handler
    
    SaveCurrentFile False
    
    Exit Sub
    
err_handler:

    ShowCurrentError
    
End Sub

Private Sub mnuSaveAs_Click()

    On Error GoTo err_handler

    SaveCurrentFile True
    
    Exit Sub
    
err_handler:

    ShowCurrentError
    
End Sub

Private Sub mnuSelectAll_Click()

    On Error GoTo err_handler

    ctlTextBox.SelectAll
    
    Exit Sub
    
err_handler:

    ShowCurrentError
    
End Sub

Private Sub mnuSettings_Click()
    Dim cFrm    As frmSettings
    
    On Error GoTo err_handler
    
    Set cFrm = New frmSettings
    
    cFrm.Show vbModal
    
    Exit Sub
    
err_handler:

    ShowCurrentError
    
End Sub
