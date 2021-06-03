VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4620
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4620
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   480
      Left            =   420
      Picture         =   "frmAbout.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   1035
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3195
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmAbout.frm - information about program and author
' // By The trick 2021
' //

Option Explicit

Private Sub Form_Load()
    
    Set Me.Icon = frmMain.Icon
    
    lblInfo.Caption = "VECTOR-06C BASIC converter." & vbNewLine & _
                      "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & _
                      "© 2021 The trick" & vbNewLine & _
                      "This program is Freeware."
    
End Sub
