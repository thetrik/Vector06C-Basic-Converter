Attribute VB_Name = "modWinapi"
' //
' // modWinapi.bas - common procedures. Before it was module with WINAPI declarations.
' // By The trick 2021
' //

Option Explicit

' // Return number of elements in array
Public Function SACount( _
                ByVal ppSA As Long) As Long
    Dim tSA     As SAFEARRAY
    Dim pBound  As Long
    Dim tBound  As SAFEARRAYBOUND
    
    If ppSA = 0 Then Exit Function
    
    GetMem4 ByVal ppSA, ppSA
    
    If ppSA = 0 Then Exit Function
    
    memcpy tSA, ByVal ppSA, Len(tSA)
    
    pBound = ppSA + Len(tSA)
    SACount = 1
    
    Do While tSA.cDims > 0
    
        memcpy tBound, ByVal pBound, Len(tBound)
        
        SACount = SACount * tBound.cElements
        pBound = pBound + Len(tBound)
        tSA.cDims = tSA.cDims - 1
        
    Loop
    
End Function

