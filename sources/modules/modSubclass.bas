Attribute VB_Name = "modSubclass"
' //
' // modSubclass.bas - module implements windows-subclassing logic
' // By The trick 2021
' //

Option Explicit

Private Const MODULE_NAME = "modSubclass"

' // Subclass a window. Returns true if successful
Public Function SubclassWindow( _
                ByVal cObj As ISubclass) As Boolean
    SubclassWindow = SetWindowSubclass(cObj.hWnd, AddressOf SubclassWndProc, 1, ByVal ObjPtr(cObj))
End Function
                
' // Unsubclass a window. Returns true if successful
Public Function UnsubclassWindow( _
                ByVal cObj As ISubclass) As Boolean
    UnsubclassWindow = RemoveWindowSubclass(cObj.hWnd, AddressOf SubclassWndProc, 1)
End Function

' // Main subclass procedure.
Private Function SubclassWndProc( _
                 ByVal hWnd As Long, _
                 ByVal lMsg As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Long, _
                 ByVal uIdSubclass As Long, _
                 ByVal dwRefData As Long) As Long
    Dim cObject     As ISubclass
    Dim bDefCall    As Boolean
    
    On Error GoTo unsubclass
    
    If dwRefData Then
        
        vbaObjSetAddref cObject, ByVal dwRefData
        SubclassWndProc = cObject.OnWindowProc(hWnd, lMsg, wParam, lParam, bDefCall)
    
    Else
        bDefCall = True
    End If
    
    If bDefCall Then
        SubclassWndProc = DefSubclassProc(hWnd, lMsg, wParam, ByVal lParam)
    End If
    
    Exit Function
    
unsubclass:

    UnsubclassByHwnd hWnd
    
End Function

Private Sub UnsubclassByHwnd( _
            ByVal hWnd As Long)
    RemoveWindowSubclass hWnd, AddressOf SubclassWndProc, 1
End Sub

