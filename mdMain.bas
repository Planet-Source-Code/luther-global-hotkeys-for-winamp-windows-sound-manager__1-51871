Attribute VB_Name = "mdMain"
Option Explicit: DefInt I

'<<<<< DECLARACIONES API / CONSTANTES >>>>>
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
        Public Const SW_HIDE = 0
        Public Const SW_SHOWNORMAL = 1
        Public Const SW_SHOWMINIMIZED = 2
        Public Const SW_SHOWMAXIMIZED = 3
        Public Const SW_SHOWNOACTIVATE = 4
        Public Const SW_MINIMIZE = 6
        Public Const SW_SHOWMINNOACTIVE = 7
        Public Const SW_SHOWNA = 8
        Public Const SW_RESTORE = 9
        Public Const SW_SHOWDEFAULT = 10
        
    Public Const cSndVol32 = "C:\WINDOWS\System32\sndvol32.exe"
    
'<<<< VARIABLES PARA LAS HOTKEYS>>>>
'    <<<<< HOTKEYS PARA EL WINAMP >>>>>
        Public iPrev, iPlay, iPause, iStop, iNext
        Public iToggleMainWnd, iTogglePlaylist
        Public iRaiseVolume, iLowerVolume, iFastFwd, iRewind
        Public iShuffle, iRepeat
        
'    <<<<< OTRAS HOTKEYS >>>>>
        Public iShowForm
        Public iOpenKLite, iOpenSndVol32

Public Function WindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If wMsg = WM_HOTKEY And frmMain.chkHotKeys Then
        

        Select Case wParam
        
        '<<<<< HOTKEYS PARA EL WINAMP >>>>>
            Case iPrev: WinampCommand (wmpPrevTrack)
            Case iPlay: WinampCommand (wmpPlay)
            Case iPause: WinampCommand (wmpTogglePause)
            Case iStop: WinampCommand (wmpStop)
            Case iNext: WinampCommand (wmpNextTrack)
                    
            Case iToggleMainWnd: WinampCommand (wmpToggleMainWindow)
            Case iTogglePlaylist: WinampCommand (wmpTogglePlaylist)
                    
            Case iRaiseVolume: WinampCommand (wmpRaiseVolume)
            Case iLowerVolume: WinampCommand (wmpLowerVolume)
            Case iFastFwd: WinampCommand (wmpFastFwd)
            Case iRewind: WinampCommand (wmpRewind)
                    
            Case iShuffle: WinampCommand (wmpToggleShuffle)
            Case iRepeat: WinampCommand (wmpToggleRepeat)
        
        '<<<<< OTRAS HOTKEYS >>>>>
            Case iShowForm: frmMain.Show
            Case iOpenSndVol32: ShellExecute frmMain.hWnd, "Open", cSndVol32, vbNull, "C:\WINDOWS\System32", SW_SHOWNORMAL
            
            
        End Select

    End If
    WindowProc = CallWindowProc(lPrevWindowProc, hWnd, wMsg, wParam, lParam)
End Function
