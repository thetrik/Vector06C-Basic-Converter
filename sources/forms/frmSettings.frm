VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3435
   Icon            =   "frmSettings.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1680
      TabIndex        =   5
      Top             =   900
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   360
      TabIndex        =   4
      Top             =   900
      Width           =   1275
   End
   Begin VB.ComboBox cboTheme 
      Height          =   315
      ItemData        =   "frmSettings.frx":000C
      Left            =   960
      List            =   "frmSettings.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2235
   End
   Begin VB.ComboBox cboFontSize 
      Height          =   315
      ItemData        =   "frmSettings.frx":0032
      Left            =   960
      List            =   "frmSettings.frx":003C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   2235
   End
   Begin VB.Label lblLabel 
      Caption         =   "Theme:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   795
   End
   Begin VB.Label lblLabel 
      Caption         =   "Font size:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmSettings.frm - settings form
' // By The trick 2021
' //

Option Explicit

Private Const MODULE_NAME = "frmSettings"

Private m_bApply    As Boolean  ' // If true - save settings

Private Sub cboFontSize_Click()
    Const PROC_NAME = "cboFontSize_Click", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim eSize   As eFontSize
    
    On Error GoTo error_handler
    
    Select Case cboFontSize.ListIndex
    Case 0: eSize = FS_SMALL
    Case 1: eSize = FS_LARGE
    End Select
    
    frmMain.SetTextboxFontSize eSize
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub cboTheme_Click()
    Const PROC_NAME = "cboTheme_Click", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim eTheme  As eTheme
    
    On Error GoTo error_handler
    
    Select Case cboTheme.ListIndex
    Case 0: eTheme = T_WIN
    Case 1: eTheme = T_DARK
    Case 2: eTheme = T_SOFT
    End Select
    
    frmMain.SetTextboxTheme eTheme
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    m_bApply = True
    Unload Me
End Sub

Private Sub Form_Load()
    Const PROC_NAME = "Form_Load", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim lIndex  As Long
    
    On Error GoTo error_handler
    
    Set Me.Icon = frmMain.Icon
    
    Select Case Settings(SET_THEME)
    Case T_WIN:     lIndex = 0
    Case T_DARK:    lIndex = 1
    Case T_SOFT:    lIndex = 2
    End Select
    
    cboTheme.ListIndex = lIndex
    
    Select Case Settings(SET_FONT_SIZE)
    Case FS_SMALL:  lIndex = 0
    Case FS_LARGE:  lIndex = 1
    End Select
    
    cboFontSize.ListIndex = lIndex
    
    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)
    Const PROC_NAME = "Form_Unload", FULL_PROC_NAME = MODULE_NAME & "::" & PROC_NAME
    
    Dim eTheme  As eTheme
    Dim eSize   As eFontSize
    
    On Error GoTo error_handler
    
    If m_bApply Then
    
        Select Case cboTheme.ListIndex
        Case 0: eTheme = T_WIN
        Case 1: eTheme = T_DARK
        Case 2: eTheme = T_SOFT
        End Select
    
        Settings(SET_THEME) = eTheme
        
        Select Case cboFontSize.ListIndex
        Case 0: eSize = FS_SMALL
        Case 1: eSize = FS_LARGE
        End Select
        
        Settings(SET_FONT_SIZE) = eSize
    
    Else
        
        frmMain.SetTextboxTheme Settings(SET_THEME)
        frmMain.SetTextboxFontSize Settings(SET_FONT_SIZE)
        
    End If

    Exit Sub
    
error_handler:
    
    ShowCurrentError
    
End Sub
