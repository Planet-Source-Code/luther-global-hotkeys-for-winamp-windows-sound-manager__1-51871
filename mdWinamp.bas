Attribute VB_Name = "mdWinamp"
Option Explicit

'<<<<< APIs >>>>>
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'<<<<< CONSTANTES >>>>>
    '<<<<< COMANDOS DE WINAMP >>>>
        Public Const wmpPrevTrack = 40044
        Public Const wmpNextTrack = 40048
        Public Const wmpPlay = 40045
        Public Const wmpTogglePause = 40046
        Public Const wmpStop = 40047
        
        Public Const wmpFadeoutAndStop = 40147
        Public Const wmpStopAfterCurrent = 40157
        
        Public Const wmpFastFwd = 40148
        Public Const wmpRewind = 40144
        
        Public Const wmpGoTo_PlaylistSt = 40154
        Public Const wmpGoTo_PlaylistEnd = 40158
        
        Public Const wmpOpenFileDialog = 40029
        Public Const wmpOpenURLDialog = 40155
        Public Const wmpShowFileInfo = 40188
        
        Public Const wmpSetTimeDisplay_Elapsed = 40037
        Public Const wmpSetTimeDisplay_Remaining = 40038
        
        Public Const wmpTogglePreferences = 40012
        
        Public Const wmpShowVisualizationOptions = 40190
        Public Const wmpShowVisualizationPlugInOptions = 40191
        Public Const wmpExecuteCurrentPlugIn = 40192
        
        Public Const wmpToggleAboutBox = 40041
        
        Public Const wmpToggleTitleAutoscrolling = 40189
        Public Const wmpToggleAlwaysOnTop = 40019
        Public Const wmpToggleWindowshade = 40064
        Public Const wmpTogglePlaylistWindowshade = 40266
        Public Const wmpToggleDoublesize = 40165
        
        Public Const wmpToggleEQ = 40036
        Public Const wmpTogglePlaylist = 40040
        Public Const wmpToggleMainWindow = 40258
        Public Const wmpToggleMiniBrowser = 40298
        
        Public Const wmpToggleEasymove = 40186
        
        Public Const wmpRaiseVolume = 40058
        Public Const wmpLowerVolume = 40059
        
        Public Const wmpToggleRepeat = 40022
        Public Const wmpToggleShuffle = 40023
        
        Public Const wmpShowJumpToTime = 40193
        Public Const wmpShowJumpToFile = 40194
        
        Public Const wmpShowSkinSelector = 40219
        Public Const wmpConfigureCurrentPlugIn = 40221
        Public Const wmpReload_the_current_skin = 40291
        
        Public Const wmpCloseWinamp = 40001
        
        Public Const wmp10TracksBack = 40197
        
        Public Const wmpEditBookmarks = 40320
        Public Const wmpBookmarkCurrentTrack = 40321
        
        Public Const wmpPlayCD = 40323
        
        Public Const wmpLoadEQPreset = 40253
        Public Const wmpShowLoadPreset = 40172
        Public Const wmpShowAutoLoadPresets = 40173
        Public Const wmpLoadDefaultPreset = 40174
        
        Public Const wmpSavePreset = 40254
        Public Const wmpShowSavePreset = 40175
        Public Const wmpShowAutoSavePreset = 40176
        
        Public Const wmShowDeletePreset = 40178
        Public Const wmpShowDeletAutoPreset = 40180
    '<<<<< OTRAS >>>>>
        Public Const cWMPWINDOW As String = "Winamp v1.x"
        Public Const cWMPCOMMAND = &H111
    
    
Public Function WinampRunning() As Boolean
    WinampRunning = (FindWindow(cWMPWINDOW, vbNullString) <> 0)
End Function

Public Function WinampCommand(ByVal lCommand As Long) As Long
    WinampCommand = SendMessage(FindWindow(cWMPWINDOW, vbNullString), cWMPCOMMAND, lCommand, 0)
End Function
