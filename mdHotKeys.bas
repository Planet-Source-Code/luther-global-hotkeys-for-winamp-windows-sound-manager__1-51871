Attribute VB_Name = "mdHotKeys"
Option Explicit: DefLng L: DefInt I: DefStr S

Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer

Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MOD_WIN = &H8

Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)

Public lPrevWindowProc

Public Sub NewHotKey(ByVal hWnd As Long, ByRef ID, ByVal lModifiers, _
                     ByVal lKey, ByVal sAtomName)
    
    'ATENCIÓN: El integer ID está declarado ByRef para poder reconocer
    '          la hotkey en el WndProc. Usar un integer público.
    
    ID = GlobalAddAtom(sAtomName)
    RegisterHotKey hWnd, ID, lModifiers, lKey
End Sub

Public Sub DeleteHotKey(iHotKeyID)
    UnregisterHotKey frmMain.hWnd, iHotKeyID
    GlobalDeleteAtom iHotKeyID
End Sub

Public Sub HookApp(ByVal hWnd As Long)
    lPrevWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookApp(ByVal hWnd As Long)
    SetWindowLong hWnd, GWL_WNDPROC, lPrevWindowProc
End Sub

'Public Function WindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Quitar los comentarios y agregar código.
'    WindowProc = CallWindowProc(lPrevWindowProc, hWnd, wMsg, wParam, lParam)
'End Function
