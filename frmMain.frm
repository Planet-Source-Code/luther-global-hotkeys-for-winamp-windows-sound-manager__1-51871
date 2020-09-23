VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HotKey Control"
   ClientHeight    =   450
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminate Program"
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   1575
   End
   Begin VB.CheckBox chkHotKeys 
      Caption         =   "Enable HotKeys"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TerminateApp As Boolean

Private Sub cmdTerminate_Click()
    TerminateApp = True
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Hide
    Me.Top = Screen.Height - Me.Height - 450
    Me.Left = Screen.Width - Me.Width
    
    ReDim HotKeyArray(1 To 1)
    
    '<<<<< HOTKEYS PARA EL WINAMP >>>>>
        NewHotKey Me.hWnd, iPlay, MOD_WIN, vbKeyX, "wmpCmdPlay"
        NewHotKey Me.hWnd, iPause, MOD_WIN, vbKeyC, "wmpCmdPause"
        NewHotKey Me.hWnd, iStop, MOD_WIN, vbKeyV, "wmpCmdStop"
        NewHotKey Me.hWnd, iPrev, MOD_WIN + MOD_CONTROL, vbKeyZ, "wmpCmdPrev"
        NewHotKey Me.hWnd, iNext, MOD_WIN + MOD_ALT, vbKeyZ, "wmpCmdNext"
        
        NewHotKey Me.hWnd, iShuffle, MOD_WIN, vbKeyS, "wmpCmdShuffle"
        NewHotKey Me.hWnd, iRepeat, MOD_WIN, vbKeyA, "wmpCmdRepeat"
        
        NewHotKey Me.hWnd, iToggleMainWnd, MOD_WIN + MOD_CONTROL, vbKeyX, "wmpCmdToggleMainWindow"
        NewHotKey Me.hWnd, iTogglePlaylist, MOD_WIN + MOD_CONTROL, vbKeyC, "wmpCmdTogglePlaylist"
        
        NewHotKey Me.hWnd, iRaiseVolume, MOD_WIN, vbKeyUp, "wmpCmdRaiseVolume"
        NewHotKey Me.hWnd, iLowerVolume, MOD_WIN, vbKeyDown, "wmpCmdLowerVolume"
        
        NewHotKey Me.hWnd, iFastFwd, MOD_WIN, vbKeyRight, "wmpCmdFastFwd"
        NewHotKey Me.hWnd, iRewind, MOD_WIN, vbKeyLeft, "wmpCmdRewind"
        
    '<<<<< HOTKEYS DE PROGRAMAS >>>>>
        NewHotKey Me.hWnd, iOpenSndVol32, MOD_CONTROL + MOD_WIN, vbKeyV, "OpenSndVol32"
    
    '<<<<< OTRAS HOTKEYS >>>>>
        NewHotKey Me.hWnd, iShowForm, MOD_CONTROL, vbKeyF12, "ShowForm"
    
    '<<<<< HOOK >>>>>
        HookApp Me.hWnd
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not TerminateApp Then Cancel = 1: Me.Hide: Exit Sub
    
    DeleteHotKey iPlay
    DeleteHotKey iPause
    DeleteHotKey iStop
    DeleteHotKey iPrev
    DeleteHotKey iNext
    
    DeleteHotKey iShuffle
    DeleteHotKey iRepeat
    
    DeleteHotKey iToggleMainWnd
    DeleteHotKey iTogglePlaylist
    
    DeleteHotKey iRaiseVolume
    DeleteHotKey iLowerVolume
    
    DeleteHotKey iFastFwd
    DeleteHotKey iRewind
    
    DeleteHotKey iOpenSndVol32
    DeleteHotKey iShowForm
    
    End
End Sub
