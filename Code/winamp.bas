Attribute VB_Name = "modWinamp"
'WinampMod - winamp.bas
'By MidTerror - midterror@hotmail.com
'Feel free to mail comments, bugs, advice
'I'd appreciate it if you gave me some credit
'If you make a program out of this. If not
'then it's ok, but I wouldn't mind seeing the
'program, so feel free to send me programs you
'make with this.

Option Explicit

'All the Declarations
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageCDS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As COPYDATASTRUCT) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_USER = &H400
Public Const WM_WA_IPC = WM_USER
Public Const WM_COPYDATA = &H4A
Public Const WM_COMMAND = &H111
Public Const WM_CHAR = &H102

'Registry Info
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Const WINAMP_REG_KEY = "WinAmp.File\shell\play\command"
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Const KEY_QUERY_VALUE = &H1

Public hwnd_winamp As Long
Public Const WM_GETTEXT = &HD
Public Const vbKeyShift = &H10
Public Const vbKeyCtrl = &H11
Public Const vbKeyAlt = &H12
Public Const IPC_DELETE = 101
Public Const IPC_ISPLAYING = 104
Public Const IPC_GETOUTPUTTIME = 105
Public Const IPC_JUMPTOTIME = 106
Public Const IPC_WRITEPLAYLIST = 120
Public Const IPC_SETPLAYLISTPOS = 121
Public Const IPC_SETVOLUME = 122
Public Const IPC_SETPANNING = 123
Public Const IPC_GETLISTLENGTH = 124
Public Const IPC_SETSKIN = 200
Public Const IPC_GETSKIN = 201
Public Const IPC_GETLISTPOS = 125
Public Const IPC_GETINFO = 126
Public Const IPC_GETEQDATA = 127
Public Const IPC_PLAYFILE = 100
Public Const IPC_CHDIR = 103
Public Const WINAMP_OPTIONS_EQ = 40036
Public Const WINAMP_OPTIONS_PLEDIT = 40040
Public Const WINAMP_VOLUMEUP = 40058
Public Const WINAMP_VOLUMEDOWN = 40059
Public Const WINAMP_FFWD5S = 40060
Public Const WINAMP_REW5S = 40061
Public Const WINAMP_BUTTON1 = 40044
Public Const WINAMP_BUTTON2 = 40045
Public Const WINAMP_BUTTON3 = 40046
Public Const WINAMP_BUTTON4 = 40047
Public Const WINAMP_BUTTON5 = 40048
Public Const WINAMP_BUTTON1_SHIFT = 40144
Public Const WINAMP_BUTTON4_SHIFT = 40147
Public Const WINAMP_BUTTON5_SHIFT = 40148
Public Const WINAMP_BUTTON1_CTRL = 40154
Public Const WINAMP_BUTTON2_CTRL = 40155
Public Const WINAMP_BUTTON5_CTRL = 40158
Public Const WINAMP_FILE_PLAY = 40029
Public Const WINAMP_OPTIONS_PREFS = 40012
Public Const WINAMP_OPTIONS_AOT = 40019
Public Const WINAMP_HELP_ABOUT = 40041

Public Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As String
End Type

'Song Info
Public Type Mp3InfoType
        Title As String
        Artist As String
        Album As String
        Year As Integer
        Genre As String
        Comment As String
End Type

Public Mp3Info As Mp3InfoType

'Gets path/filename
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

'TypeText Declares
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Const KEYEVENTF_KEYUP = &H2

Public Function TypeText(TextToType As String, WindowToTypeIn As Long) As Long
'This is a function I made to send text
'To a certain application using modules
'Has nothing to do with winamp, you can skip it, I use
'it in a different function
    Dim mVK As Long
    Dim mScan As Long
    Dim a As Integer
    Dim CurrentForeground As Long
    Dim GiveUpCount As Integer
    Dim ShiftDown As Boolean, AltDown As Boolean, ControlDown As Boolean
    
    If TextToType = "" Then Exit Function
    
    CurrentForeground = GetForegroundWindow()
    
    For a = 1 To Len(TextToType)
        
        mVK = VkKeyScan(Asc(Mid(TextToType, a, 1)))
        mScan = MapVirtualKey(mVK, 0)
        
        ShiftDown = (mVK And &H100)
        ControlDown = (mVK And &H200)
        AltDown = (mVK And &H400)
        
        mVK = mVK And &HFF
        
        GiveUpCount = 0
        
        Do While GetForegroundWindow() <> WindowToTypeIn And GiveUpCount < 20
            GiveUpCount = GiveUpCount + 1
            SetForegroundWindow WindowToTypeIn
            DoEvents
        Loop
        
        If GetForegroundWindow() <> WindowToTypeIn Then TypeText = 0: Exit Function
        
        If ShiftDown Then keybd_event &H10, 0, 0, 0
        If ControlDown And &H200 Then keybd_event &H11, 0, 0, 0
        If AltDown And &H400 Then keybd_event &H12, 0, 0, 0
        
        keybd_event mVK, mScan, 0, 0
        
        If ShiftDown Then keybd_event &H10, 0, KEYEVENTF_KEYUP, 0
        If ControlDown Then keybd_event &H11, 0, KEYEVENTF_KEYUP, 0
        If AltDown Then keybd_event &H12, 0, KEYEVENTF_KEYUP, 0
        
    Next a
    
    SetForegroundWindow CurrentForeground
    
    TypeText = 1
    
End Function

Public Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
'I got this from Microsoft's page and editted it a bit
'Has nothing to do with winamp, you can skip it, I use
'it in a different function
    Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long, v$, r As Long
    RetVal$ = ""
    r = RegOpenKeyEx(hInKey, subkey$, 0, KEY_QUERY_VALUE, hSubKey)
    If r <> 0 Then Exit Function
    SZ = 256
    v$ = String$(SZ, 0)
    r = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)
    If r = 0 And dwType = 1 Then
        RetVal$ = Left(v$, SZ - 1)
    Else
        RetVal$ = ""
    End If


    If hInKey = 0 Then r = RegCloseKey(hSubKey)
    RegGetString$ = RetVal$
End Function

Public Function FindWinamp() As Long
'Find winamp window
'Returns 1 if winamp is open, 0 if not
    hwnd_winamp = FindWindow("Winamp v1.x", vbNullString)
    If hwnd_winamp Then FindWinamp = 1 Else FindWinamp = 0
End Function

Public Function DeletePlayList() As Long
'Clears the play list
    DeletePlayList = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_DELETE)
End Function
Public Function IsPlaying() As Long
'Returns:
'1 If playing
'3 if paused
'0 if stopped
    IsPlaying = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_ISPLAYING)
End Function

Public Function GetCurrentSongPosition() As Double
'Finds the current song position in milliseconds
    GetCurrentSongPosition = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETOUTPUTTIME)
End Function

Public Function GetSongLength() As Long
'Finds the song length in Seconds
    GetSongLength = SendMessage(hwnd_winamp, WM_WA_IPC, 1, IPC_GETOUTPUTTIME)
End Function

Public Function SetCurrentSongPosition(Optional Seconds As Long, Optional Ms As Long)
'Sets the current position in the song
'Returns:
'0 if success
'1 if eof
'-1 if not playing
    SetCurrentSongPosition = SendMessage(hwnd_winamp, WM_WA_IPC, (Seconds * 1000 + Ms), IPC_JUMPTOTIME)
End Function


Public Function WritePlayList() As Long
'Writes the current playlist to C:\WINAMP_DIR\Winamp.m3u
'And then finds the play position
'Now obsolete, but good for old version of winamp
'Look at GetPlayListPosition
    WritePlayList = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_WRITEPLAYLIST)
End Function

Public Function SetPlayListPosition(Position As Integer) As Long
'Sets which song to play (0 being first)
    SetPlayListPosition = SendMessage(hwnd_winamp, WM_WA_IPC, Position, IPC_SETPLAYLISTPOS)
End Function

Public Function SetVolume(Volume As Integer) As Long
'Sets the volume (Volume must be between 0 - 255)
    SetVolume = SendMessage(hwnd_winamp, WM_WA_IPC, Volume, IPC_SETVOLUME)
End Function

Public Function SetPanning(PanPosition As Integer) As Long
'Sets the panning (PanPosition must be between 0 - 255)
    SetPanning = SendMessage(hwnd_winamp, WM_WA_IPC, PanPosition, IPC_SETPANNING)
End Function

Public Function GetPlayListLength() As Long
'Gets amount of songs in play list
    GetPlayListLength = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETLISTLENGTH)
End Function


Public Function GetPlayListPosition() As Long
'Returns which song its playing in the playlist
'0 being first
    GetPlayListPosition = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETLISTPOS)
End Function

Public Function GetSamplerate() As Long
'Gets the samplerate
    GetSamplerate = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETINFO)
End Function

Public Function GetBitrate() As Long
'Gets the bitrate
    GetBitrate = SendMessage(hwnd_winamp, WM_WA_IPC, 1, IPC_GETINFO)
End Function

Public Function GetChannels() As Long
'Gets the channel
    GetChannels = SendMessage(hwnd_winamp, WM_WA_IPC, 2, IPC_GETINFO)
End Function

Public Function GetEQBandData(BandNumber As Integer) As Long
'Get each EQ banddata (0 being the first, 9 being last)
'Returns 0 - 255
    If BandNumber > 9 Then Exit Function
    GetEQBandData = SendMessage(hwnd_winamp, WM_WA_IPC, BandNumber, IPC_GETEQDATA)
End Function

Public Function GetEQPreampValue() As Long
'Gets the preamp value (Between 0 - 255)
    GetEQPreampValue = SendMessage(hwnd_winamp, WM_WA_IPC, 10, IPC_GETEQDATA)
End Function

Public Function GetEQEnabled()
'1 if EQ is enabled
'0 if it isn't
    GetEQEnabled = SendMessage(hwnd_winamp, WM_WA_IPC, 11, IPC_GETEQDATA)
End Function

Public Function GetEQAutoLoad()
'1 if EQ is autoloaded
'0 if it isn't
    GetEQAutoLoad = SendMessage(hwnd_winamp, WM_WA_IPC, 12, IPC_GETEQDATA)
End Function

Public Function PlayFile(FileToPlay As String) As Long
'Adds FileToPlay to the play list
    Dim CDS As COPYDATASTRUCT
    CDS.dwData = IPC_PLAYFILE
    CDS.lpData = FileToPlay
    CDS.cbData = Len(FileToPlay) + 1
    PlayFile = SendMessageCDS(hwnd_winamp, WM_COPYDATA, 0, CDS)
End Function

Public Function ChangeDirectory(Directory As String) As Long
'Changes directory
    Dim CDS As COPYDATASTRUCT
    CDS.dwData = IPC_CHDIR
    CDS.lpData = Directory
    CDS.cbData = Len(Directory) + 1
    ChangeDirectory = SendMessageCDS(hwnd_winamp, WM_COPYDATA, 0, CDS)
End Function

Public Function ToggleEQWindow() As Long
'Turns on or off the EQ window
    ToggleEQWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_EQ, 0)
End Function

Public Function TogglePlayListWindow() As Long
'Turns on or off play list window
    TogglePlayListWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_PLEDIT, 0)
End Function

Public Function VolumeUp() As Long
'Raises the volume a tiny bit
    VolumeUp = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_VOLUMEUP, 0)
End Function
Public Function VolumeDown() As Long
'Sets the volume down a tiny bit
    VolumeDown = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_VOLUMEDOWN, 0)
End Function

Public Function Rewind() As Long
'Rewinds by 5 seconds
    Rewind = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_REW5S, 0)
End Function

Public Function FastForward() As Long
'Fast forwards by 5 seconds
    FastForward = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_FFWD5S, 0)
End Function

Public Function PreviousSong() As Long
'Plays the previous song
    PreviousSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON1, 0)
End Function

Public Function PlaySong() As Long
'Plays the current song
    PlaySong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON2, 0)
End Function

Public Function PauseSong() As Long
'Pauses playing
    PauseSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON3, 0)
End Function
Public Function StopSong() As Long
'Stops playing
    StopSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON4, 0)
End Function

Public Function NextSong() As Long
'Plays the next song in the playlist
    NextSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON5, 0)
End Function

Public Function FadeStop() As Long
'slowly fades away until it stops
    FadeStop = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON4_SHIFT, 0)
End Function

Public Function Back10Songs() As Long
'Goes to the first song in the play list
    Back10Songs = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON1_CTRL, 0)
End Function

Public Function Forward10Songs() As Long
'Goes to the last song in the play list
    Forward10Songs = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON5_CTRL, 0)
End Function
Public Function OpenLocation() As Long
'Shows Open Location Dialog
    OpenLocation = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON2_CTRL, 0)
End Function
Public Function LoadFile() As Long
'Shows Load a file dialog
    LoadFile = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_FILE_PLAY, 0)
End Function
Public Function ShowPreferences() As Long
'Shows Preferences Dialog
    ShowPreferences = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_PREFS, 0)
End Function

Public Function ToggleAlwaysOnTop() As Long
'Turns Always On Top On and Off
    ToggleAlwaysOnTop = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_AOT, 0)
End Function

Public Function ShowAbout() As Long
'Shows About Box
    ShowAbout = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_HELP_ABOUT, 0)
End Function




Public Function ToggleRepeat() As Long
'Turns On/Off the repeat songs
    ToggleRepeat = TypeText("r", hwnd_winamp)
End Function

Public Function ToggleShuffle() As Long
'Turns On/Off the shuffle songs
    ToggleShuffle = TypeText("s", hwnd_winamp)
End Function

Public Function ToggleWindowShade() As Long
'Turns On/Off Window Shade mode
    keybd_event vbKeyCtrl, 0, 0, 0
        ToggleWindowShade = TypeText("w", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ToggleDoubleSize() As Long
'Turns on/off doublesize
    keybd_event vbKeyCtrl, 0, 0, 0
        ToggleDoubleSize = TypeText("d", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ToggleEasyMove() As Long
'turns on/off easy move
    keybd_event vbKeyCtrl, 0, 0, 0
        ToggleEasyMove = TypeText("r", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ToggleTimeDisplay() As Long
'Changes type of time display
    keybd_event vbKeyCtrl, 0, 0, 0
        ToggleTimeDisplay = TypeText("t", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ToggleMainWindow() As Long
'Hides/Shows winamp
    keybd_event vbKeyAlt, 0, 0, 0
        ToggleMainWindow = TypeText("w", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ToggleMiniBrowser() As Long
'Hides/Shows Mini Browser
    keybd_event vbKeyAlt, 0, 0, 0
        ToggleMiniBrowser = TypeText("t", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ShowSkinBrowser() As Long
'Shows Skin Browser
    keybd_event vbKeyAlt, 0, 0, 0
        ShowSkinBrowser = TypeText("s", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function
Public Function ShowVisualOptions() As Long
'Shows Visual Options
    keybd_event vbKeyAlt, 0, 0, 0
        ShowVisualOptions = TypeText("o", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ShowConfigureVisualPlugin() As Long
'Shows Configuration for current visual plugin
    keybd_event vbKeyAlt, 0, 0, 0
        ShowConfigureVisualPlugin = TypeText("k", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ToggleVisualPlugin() As Long
'Shows/Hides visual plugin
    keybd_event vbKeyAlt, 0, 0, 0
        ToggleVisualPlugin = TypeText("K", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function ShowVisualPluginsSelect() As Long
'Shows visual plugins selection
    keybd_event vbKeyCtrl, 0, 0, 0
        ShowVisualPluginsSelect = TypeText("k", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function


Public Function StopAfterCurrentSong() As Long
'Stop playing after current song
    keybd_event vbKeyCtrl, 0, 0, 0
        StopAfterCurrentSong = TypeText("v", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function OpenDirectory() As Long
'Shows Open Directory dialog
    OpenDirectory = TypeText("L", hwnd_winamp)
End Function

Public Function ShowInfoBox() As Long
'Shows info box for current song
    keybd_event vbKeyAlt, 0, 0, 0
        ShowInfoBox = TypeText("3", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function GetMp3Info() As Long
'Finds all the info about the current mp3 and sets
'Mp3Info to the info
'Mp3Info.Title = title
'Mp3Info.Artist = artist and so on
    
    Dim hwnd_InfoBox As Long
    Dim hwnd_TmpText As Long
    Dim TmpText As String * 35
    Dim TextLen As Long
    
    ShowInfoBox
    DoEvents
    
    Do While hwnd_InfoBox = 0
        DoEvents
        hwnd_InfoBox = FindWindow("#32770", "MPEG file info box + ID3 tag editor")
    Loop
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, 0, "Edit", vbNullString)
    TextLen = SendMessageS(hwnd_TmpText, WM_GETTEXT, Len(TmpText), TmpText)
    Mp3Info.Title = Left(TmpText, TextLen)
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, hwnd_TmpText, "Edit", vbNullString)
    TextLen = SendMessageS(hwnd_TmpText, WM_GETTEXT, Len(TmpText), TmpText)
    Mp3Info.Artist = Left(TmpText, TextLen)
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, hwnd_TmpText, "Edit", vbNullString)
    TextLen = SendMessageS(hwnd_TmpText, WM_GETTEXT, Len(TmpText), TmpText)
    Mp3Info.Album = Left(TmpText, TextLen)
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, hwnd_TmpText, "Edit", vbNullString)
    TextLen = SendMessageS(hwnd_TmpText, WM_GETTEXT, Len(TmpText), TmpText)
    Mp3Info.Year = Val(Left(TmpText, TextLen))
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, hwnd_TmpText, "Edit", vbNullString)
    TextLen = SendMessageS(hwnd_TmpText, WM_GETTEXT, Len(TmpText), TmpText)
    Mp3Info.Comment = Left(TmpText, TextLen)
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, 0, "ComboBox", vbNullString)
    TextLen = SendMessageS(hwnd_TmpText, WM_GETTEXT, Len(TmpText), TmpText)
    Mp3Info.Genre = Left(TmpText, TextLen)
    
    DoEvents
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, 0, "Button", "Cancel")
    TypeText Chr(13), hwnd_TmpText
End Function

Public Function GetWinampPath() As String
'Finds the path of winamp
    
    Dim WinampPath As String
    
    WinampPath = RegGetString(HKEY_CLASSES_ROOT, WINAMP_REG_KEY, "")
    If Len(WinampPath) < 8 Then GetWinampPath = "": Exit Function
    WinampPath = Mid(WinampPath, 2, Len(WinampPath) - 7)
    GetWinampPath = WinampPath
End Function

Public Function GetCurrentSongPath() As String
'Finds the path of the song currently playing
Dim CurrentPosition As Integer
Dim PathOfWinamp As String
Dim CurrentSongPath As String
Dim a As Integer

    CurrentPosition = WritePlayList()
    If CurrentPosition = -1 Then Exit Function
    PathOfWinamp = GetWinampPath
    If PathOfWinamp = "" Then Exit Function
    
    a = 1
    Do While InStr(a + 1, PathOfWinamp, "\")
        a = a + 1
    Loop
    PathOfWinamp = Left(PathOfWinamp, a)
    If FindWinamp = 0 Then Exit Function
    If WritePlayList = -1 Then Exit Function
    
    Open PathOfWinamp & "WINAMP.m3u" For Input As #1
    Line Input #1, CurrentSongPath
    For a = 1 To (CurrentPosition + 1)
        Line Input #1, CurrentSongPath
        Line Input #1, CurrentSongPath
    Next a
    Close #1
    GetCurrentSongPath = CurrentSongPath
    
End Function

Public Function GetPathOfSongInPlayList(PlayListPosition As Integer)
'Finds the path of the song in the playlist (0 being first)
Dim PathOfWinamp As String
Dim SongPath As String
Dim a As Integer


    If PlayListPosition > GetPlayListLength() Then Exit Function
    PathOfWinamp = GetWinampPath
    If PathOfWinamp = "" Then Exit Function
    
    a = 1
    Do While InStr(a + 1, PathOfWinamp, "\")
        a = a + 1
    Loop
    PathOfWinamp = Left(PathOfWinamp, a)
    If FindWinamp = 0 Then Exit Function
    If WritePlayList = -1 Then Exit Function
    
    Open PathOfWinamp & "WINAMP.m3u" For Input As #1
    Line Input #1, SongPath
    For a = 1 To (PlayListPosition + 1)
        Line Input #1, SongPath
        Line Input #1, SongPath
    Next a
    Close #1
    
    If SongPath = "#EXTM3U" Then SongPath = ""
    
    GetPathOfSongInPlayList = SongPath
    
End Function
