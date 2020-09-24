Attribute VB_Name = "NumbersOnly"
Option Explicit

Public OldWindowProc As Long
Public IPOldWindowProc(0 To 3) As Long
Public SMOldWindowProc(0 To 3) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long


Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000
Public Const GWL_WNDPROC = (-4)
Private Const WM_CONTEXTMENU = &H7B
Private Const WM_PASTE = &H302
' *********************************************
' Pass along all messages except the one that
' makes the context menu appear and paste.
' *********************************************
Public Function NewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If (msg <> WM_PASTE) And (msg <> WM_CONTEXTMENU) Then
    NewWindowProc = CallWindowProc( _
        OldWindowProc, hWnd, msg, wParam, _
        lParam)
  End If
End Function


