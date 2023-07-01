VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrickSound2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // clsTrickSound2.cls - class for sound capture and playback
' // Version 2.1
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Private Enum MMRESULT
    MMSYSERR_NOERROR = 0
    MMSYSERR_ERROR = 1
    MMSYSERR_BADDEVICEID = 2
    MMSYSERR_NOTENABLED = 3
    MMSYSERR_ALLOCATED = 4
    MMSYSERR_INVALHANDLE = 5
    MMSYSERR_NODRIVER = 6
    MMSYSERR_NOMEM = 7
    MMSYSERR_NOTSUPPORTED = 8
    MMSYSERR_BADERRNUM = 9
    MMSYSERR_INVALFLAG = 10
    MMSYSERR_INVALPARAM = 11
    MMSYSERR_HANDLEBUSY = 12
    MMSYSERR_INVALIDALIAS = 13
    MMSYSERR_BADDB = 14
    MMSYSERR_KEYNOTFOUND = 15
    MMSYSERR_READERROR = 16
    MMSYSERR_WRITEERROR = 17
    MMSYSERR_DELETEERROR = 18
    MMSYSERR_VALNOTFOUND = 19
    MMSYSERR_NODRIVERCB = 20
    WAVERR_BADFORMAT = 32
    WAVERR_STILLPLAYING = 33
    WAVERR_UNPREPARED = 34
    MMRESULT_END
End Enum

Public Enum Errors
    CAPTURE_IS_ALREADY_RUNNING = vbObjectError Or (MMRESULT_END)
    INVALID_BUFFERS_COUNT
    NOT_INITIALIZE
    ERROR_UNAVAILABLE
    ERROR_OBJECT_FAILED
    ERROR_OPEN_DEVICE = vbObjectError Or (2 * &H100)
    ERROR_PREPARE_BUFFERS = vbObjectError Or (3 * &H100)
    ERROR_ADD_BUFFERS = vbObjectError Or (4 * &H100)
    ERROR_STARTUP = vbObjectError Or (5 * &H100)
    ERROR_STOP = vbObjectError Or (6 * &H100)
    ERROR_CLOSE_DEVICE = vbObjectError Or (7 * &H100)
    ERROR_RESET_DEVICE = vbObjectError Or (8 * &H100)
End Enum

Private Type WNDCLASSEX
    cbSize              As Long
    style               As Long
    lpfnwndproc         As Long
    cbClsextra          As Long
    cbWndExtra2         As Long
    hInstance           As Long
    hIcon               As Long
    hCursor             As Long
    hbrBackground       As Long
    lpszMenuName        As Long
    lpszClassName       As Long
    hIconSm             As Long
End Type

Private Type WAVEFORMATEX
    wFormatTag          As Integer
    nChannels           As Integer
    nSamplesPerSec      As Long
    nAvgBytesPerSec     As Long
    nBlockAlign         As Integer
    wBitsPerSample      As Integer
    cbSize              As Integer
End Type

Private Type WAVEINCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname(31)         As Integer
    dwFormats           As Long
    wChannels           As Integer
    wReserved1          As Integer
End Type
Private Type WAVEOUTCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname(31)         As Integer
    dwFormats           As Long
    wChannels           As Integer
    wReserved           As Integer
    dwSupport           As Long
End Type

Private Type WAVEHDR
    lpData              As Long
    dwBufferLength      As Long
    dwBytesRecorded     As Long
    dwUser              As Long
    dwFlags             As Long
    dwLoops             As Long
    lpNext              As Long
    Reserved            As Long
End Type
 
Private Type buffer
    data()              As Byte
    Header              As WAVEHDR
    Status              As Boolean
End Type

Private Type PROCESS_HEAP_ENTRY
    lpData              As Long
    cbData              As Long
    cbOverhead          As Byte
    iRegionIndex        As Byte
    wFlags              As Integer
    dwCommittedSize     As Long
    dwUnCommittedSize   As Long
    lpFirstBlock        As Long
    lpLastBlock         As Long
End Type


Private Declare Function DefWindowProc Lib "user32" _
                         Alias "DefWindowProcW" ( _
                         ByVal hwnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function HeapCreate Lib "kernel32" ( _
                         ByVal flOptions As Long, _
                         ByVal dwInitialSize As Long, _
                         ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByRef lpMem As Any) As Long
Private Declare Function HeapWalk Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByRef lpEntry As PROCESS_HEAP_ENTRY) As Long
Private Declare Function HeapLock Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function HeapUnlock Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" _
                         Alias "SetEnvironmentVariableW" ( _
                         ByVal lpName As Long, _
                         ByVal lpValue As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" _
                         Alias "GetEnvironmentVariableW" ( _
                         ByVal lpName As Long, _
                         ByVal lpBuffer As Long, _
                         ByVal nSize As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetClassInfoEx Lib "user32" _
                         Alias "GetClassInfoExW" ( _
                         ByVal hInstance As Long, _
                         ByVal lpClassName As Long, _
                         ByRef lpWndClassEx As WNDCLASSEX) As Long
Private Declare Function UnregisterClass Lib "user32" _
                         Alias "UnregisterClassW" ( _
                         ByVal lpClassName As Long, _
                         ByVal hInstance As Long) As Long
Private Declare Function RegisterClassEx Lib "user32" _
                         Alias "RegisterClassExW" ( _
                         ByRef pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function CreateWindowEx Lib "user32" _
                         Alias "CreateWindowExW" ( _
                         ByVal dwExStyle As Long, _
                         ByVal lpClassName As Long, _
                         ByVal lpWindowName As Long, _
                         ByVal dwStyle As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal nWidth As Long, _
                         ByVal nHeight As Long, _
                         ByVal hWndParent As Long, _
                         ByVal hMenu As Long, _
                         ByVal hInstance As Long, _
                         ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" ( _
                         ByVal hwnd As Long) As Long
Private Declare Function lstrlen Lib "kernel32" _
                         Alias "lstrlenW" ( _
                         ByRef lpString As Any) As Long
Private Declare Function lstrcpyn Lib "kernel32" _
                         Alias "lstrcpynW" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any, _
                         ByVal iMaxLength As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
                         Alias "SetWindowLongA" ( _
                         ByVal hwnd As Long, _
                         ByVal nIndex As Long, _
                         ByVal dwNewLong As Long) As Long
Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveInGetID Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long, _
                         ByRef lpuDeviceID As Long) As Long
Private Declare Function waveInGetDevCaps Lib "winmm.dll" _
                         Alias "waveInGetDevCapsW" ( _
                         ByVal uDeviceID As Long, _
                         ByRef lpCaps As WAVEINCAPS, _
                         ByVal uSize As Long) As Long
Private Declare Function waveInOpen Lib "winmm.dll" ( _
                         ByRef lphWaveIn As Long, _
                         ByVal uDeviceID As Long, _
                         ByRef lpFormat As WAVEFORMATEX, _
                         ByVal dwCallback As Long, _
                         ByVal dwInstance As Long, _
                         ByVal dwFlags As Long) As MMRESULT
Private Declare Function waveInPrepareHeader Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long, _
                         ByRef lpWaveInHdr As WAVEHDR, _
                         ByVal uSize As Long) As MMRESULT
Private Declare Function waveInReset Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long) As MMRESULT
Private Declare Function waveInStart Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long) As MMRESULT
Private Declare Function waveInStop Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long) As MMRESULT
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long, _
                         ByRef lpWaveInHdr As WAVEHDR, _
                         ByVal uSize As Long) As MMRESULT
Private Declare Function waveInClose Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long) As MMRESULT
Private Declare Function waveInGetErrorText Lib "winmm.dll" _
                         Alias "waveInGetErrorTextW" ( _
                         ByVal err As Long, _
                         ByVal lpText As Long, _
                         ByVal uSize As Long) As MMRESULT
Private Declare Function waveInAddBuffer Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long, _
                         ByRef lpWaveInHdr As WAVEHDR, _
                         ByVal uSize As Long) As MMRESULT
Private Declare Function waveOutGetDevCaps Lib "winmm.dll" _
                         Alias "waveOutGetDevCapsW" ( _
                         ByVal uDeviceID As Long, _
                         ByRef lpCaps As WAVEOUTCAPS, _
                         ByVal uSize As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveOutGetID Lib "winmm.dll" ( _
                         ByVal hWaveOut As Long, _
                         ByRef lpuDeviceID As Long) As Long
Private Declare Function waveOutOpen Lib "winmm.dll" ( _
                         ByRef lphWaveOut As Long, _
                         ByVal uDeviceID As Long, _
                         ByRef lpFormat As WAVEFORMATEX, _
                         ByVal dwCallback As Long, _
                         ByVal dwInstance As Long, _
                         ByVal dwFlags As Long) As MMRESULT
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" ( _
                         ByVal hWaveOut As Long, _
                         ByRef lpWaveOutHdr As WAVEHDR, _
                         ByVal uSize As Long) As MMRESULT
Private Declare Function waveOutWrite Lib "winmm.dll" ( _
                         ByVal hWaveOut As Long, _
                         ByRef lpWaveOutHdr As WAVEHDR, _
                         ByVal uSize As Long) As MMRESULT
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" ( _
                         ByVal hWaveOut As Long, _
                         ByRef lpWaveOutHdr As WAVEHDR, _
                         ByVal uSize As Long) As MMRESULT
Private Declare Function waveOutClose Lib "winmm.dll" ( _
                         ByVal hWaveOut As Long) As MMRESULT
Private Declare Function waveOutReset Lib "winmm.dll" ( _
                         ByVal hWaveOut As Long) As MMRESULT
Private Declare Function waveOutPause Lib "winmm.dll" ( _
                         ByVal hWaveOut As Long) As MMRESULT
Private Declare Function waveOutRestart Lib "winmm.dll" ( _
                         ByVal hWaveOut As Long) As MMRESULT
                         
Private Declare Sub ZeroMemory Lib "kernel32" _
                    Alias "RtlZeroMemory" ( _
                    ByRef Destination As Any, _
                    ByVal Length As Long)
Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)

Private Const SndClass                      As String = "TrickSoundClass"
Private Const HWND_MESSAGE                  As Long = -3
Private Const WAVE_MAPPER                   As Long = -1&
Private Const CALLBACK_WINDOW               As Long = &H10000
Private Const WAVE_FORMAT_PCM               As Long = 1
Private Const MM_WIM_DATA                   As Long = &H3C0
Private Const MM_WOM_DONE                   As Long = &H3BD
Private Const WNDPROCINDEX                  As Long = 20
Private Const HEAP_CREATE_ENABLE_EXECUTE    As Long = &H40000
Private Const HEAP_NO_SERIALIZE             As Long = &H1
Private Const HEAP_ZERO_MEMORY              As Long = &H8
Private Const PROCESS_HEAP_ENTRY_BUSY       As Long = &H4
Private Const GWL_WNDPROC                   As Long = (-4)

Private Init        As Boolean              ' // Whether class is initialized
Private hwnd        As Long                 ' // Handle of receiver window
Private mActive     As Boolean              ' // Whether process of capture/playback is active or not
Private mSmpCount   As Long                 ' // Size of buffer in samples
Private mFormat     As WAVEFORMATEX         ' // Format of wave
Private hWaveIn     As Long                 ' // Handle of capture device
Private hWaveOut    As Long                 ' // Handle of playback device
Private Buffers()   As buffer               ' // Buffers array
Private bufCount    As Long                 ' // Count of buffers
Private unavailable As Boolean              ' // Determine if class is unavailable
Private paused      As Boolean              ' // If pause is active
Private devCap      As Collection           ' // List of capture devices
Private devPlay     As Collection           ' // List of playback devices

Dim hHeap   As Long
Dim lpAsm   As Long

' // The event that get raised when new buffer is needed
Public Event NewData( _
             ByVal DataPtr As Long, _
             ByVal CountBytes As Long)

' // If capture/playback is active then true
Public Property Get IsActive() As Boolean
    IsActive = mActive
End Property

' // If capture/playback is initialized then true
Public Property Get IsUnavailable() As Boolean
    IsUnavailable = unavailable
End Property

' // If object initialization failed then ture
Public Property Get IsFailed() As Boolean
    IsFailed = Not Init
End Property

' // Size of buffer in second
Public Property Get BufferLengthSec() As Single
    BufferLengthSec = mSmpCount / mFormat.nSamplesPerSec
End Property

' // Size of buffer in samples
Public Property Get BufferLengthSamples() As Long
    BufferLengthSamples = mSmpCount
End Property

' // Sample rate
Public Property Get SampleRate() As Long
    SampleRate = mFormat.nSamplesPerSec
End Property

' // Bit per sample
Public Property Get BitsPerSample() As Integer
    BitsPerSample = mFormat.wBitsPerSample
End Property

' // Count of channels
Public Property Get Channels() As Integer
    Channels = mFormat.nChannels
End Property

' // Number of buffers
Public Property Get BuffersCount() As Byte
    BuffersCount = bufCount
End Property

' // Current capture device id
Public Property Get CurrentCaptureDeviceID() As Long
    If hWaveIn Then
        waveInGetID hWaveIn, CurrentCaptureDeviceID
    Else
        err.Raise 5
    End If
End Property

' // Current playback device id
Public Property Get CurrentPlaybackDeviceID() As Long
    If hWaveOut Then
        waveOutGetID hWaveOut, CurrentPlaybackDeviceID
    Else
        err.Raise 5
    End If
End Property

' // List of available capture devices
Public Property Get CaptureDevices() As Collection
    Dim devCount    As Long
    Dim caps        As WAVEINCAPS
    Dim idx         As Long
    Dim strLen      As Long
    Dim tmpStr      As String
    
    If devCap Is Nothing Then
    
        devCount = waveInGetNumDevs()
        Set devCap = New Collection
        
        For idx = 0 To devCount - 1
            waveInGetDevCaps idx, caps, Len(caps)
            strLen = lstrlen(caps.szPname(0))
            tmpStr = Space(strLen)
            lstrcpyn ByVal StrPtr(tmpStr), caps.szPname(0), strLen + 1
            devCap.Add tmpStr
        Next
    End If
    
    Set CaptureDevices = devCap
    
End Property

' // List of available playback devices
Public Property Get PlaybackDevices() As Collection
    Dim devCount    As Long
    Dim caps        As WAVEOUTCAPS
    Dim idx         As Long
    Dim strLen      As Long
    Dim tmpStr      As String
    
    If devPlay Is Nothing Then
        
        devCount = waveOutGetNumDevs()
        Set devPlay = New Collection
        
        For idx = 0 To devCount - 1
            waveOutGetDevCaps idx, caps, Len(caps)
            strLen = lstrlen(caps.szPname(0))
            tmpStr = Space(strLen)
            lstrcpyn ByVal StrPtr(tmpStr), caps.szPname(0), strLen + 1
            devPlay.Add tmpStr
        Next
        
    End If
    
    Set PlaybackDevices = devPlay
    
End Property

' // Start the capture/playback
Public Function StartProcess() As Boolean
    Dim ret As MMRESULT
    
    If mActive And Not paused Then Exit Function
    
    If Not Init Then
        err.Raise Errors.ERROR_OBJECT_FAILED
        Exit Function
    End If
    
    If Not unavailable Then
        err.Raise Errors.NOT_INITIALIZE
        Exit Function
    End If
    
    If hWaveIn Then
    
        ret = waveInStart(hWaveIn)
        If ret Then
            err.Raise ERROR_STARTUP Or ret
            Exit Function
        End If
        
    Else
    
        Dim idx As Long
        
        If paused Then
        
            ret = waveOutRestart(hWaveOut)
            
            If ret Then
                err.Raise ERROR_STARTUP Or ret
                Exit Function
            End If
            
            paused = False
            
        Else
        
            For idx = 0 To bufCount - 1
                
                RaiseEvent NewData(Buffers(idx).Header.lpData, UBound(Buffers(idx).data) + 1)
                
                ret = waveOutWrite(hWaveOut, Buffers(idx).Header, Len(Buffers(idx).Header))
                
                If ret Then
                    err.Raise ERROR_STARTUP Or ret
                    Exit Function
                End If
                
            Next
        End If

    End If
    
    StartProcess = True
    mActive = True

End Function

' // Pause playback
Public Function PauseProcess() As Boolean
    Dim ret As MMRESULT
    
    If Not Init Then
        err.Raise Errors.ERROR_OBJECT_FAILED
        Exit Function
    End If
    
    If Not unavailable Then
        err.Raise Errors.NOT_INITIALIZE
        Exit Function
    End If
    
    If Not mActive Then Exit Function
    
    If hWaveOut Then
    
        paused = True
        waveOutPause hWaveOut
        mActive = False
            
        PauseProcess = True
        
    End If
    
End Function

' // Stop playback/capture
Public Function StopProcess() As Boolean
    Dim ret As Long
    
    If Not Init Then
        err.Raise Errors.ERROR_OBJECT_FAILED
        Exit Function
    End If
    
    If Not unavailable Then
        err.Raise Errors.NOT_INITIALIZE
        Exit Function
    End If
    
    If Not mActive Then Exit Function
    
    If hWaveIn Then
        
        ret = waveInStop(hWaveIn)
        
        If ret Then
            err.Raise ERROR_STOP Or ret
            Exit Function
        End If

    Else
    
        ret = waveOutReset(hWaveOut)
        
        If ret Then
            err.Raise ERROR_STOP Or ret
            Exit Function
        End If
        
    End If
    
    mActive = False
    paused = False
    StopProcess = True
    
End Function

' // Playback initialization
Public Function InitPlayback( _
                ByVal NumOfChannels As Integer, _
                ByVal SamplesPerSec As Long, _
                ByVal BitsPerSample As Integer, _
                ByVal BufferSampleCount As Long, _
                Optional ByVal DeviceID As Long = WAVE_MAPPER, _
                Optional ByVal BuffersCount As Byte = 4) As Boolean
    Dim ret As MMRESULT
    Dim idx As Long
    
    If Not Init Then
        err.Raise Errors.ERROR_OBJECT_FAILED
        Exit Function
    End If
    
    If unavailable Then
        err.Raise Errors.ERROR_UNAVAILABLE
        Exit Function
    End If
    
    If BuffersCount < 1 Then
        err.Raise Errors.INVALID_BUFFERS_COUNT
        Exit Function
    End If
    
    unavailable = True

    With mFormat
        .cbSize = 0
        .wFormatTag = WAVE_FORMAT_PCM
        .wBitsPerSample = BitsPerSample
        .nSamplesPerSec = SamplesPerSec
        .nChannels = NumOfChannels
        .nBlockAlign = .nChannels * .wBitsPerSample \ 8
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
    End With

    mSmpCount = BufferSampleCount - (BufferSampleCount Mod mFormat.nBlockAlign)
    
    ret = waveOutOpen(hWaveOut, DeviceID, mFormat, hwnd, 0, CALLBACK_WINDOW)
    
    If ret Then
        err.Raise ERROR_OPEN_DEVICE Or ret
        Exit Function
    End If
    
    bufCount = BuffersCount
    ReDim Buffers(BuffersCount - 1)

    For idx = 0 To BuffersCount - 1
    
        With Buffers(idx)
            ReDim .data(mSmpCount * mFormat.nBlockAlign - 1)
            .Header.lpData = VarPtr(.data(0))
            .Header.dwBufferLength = UBound(.data) + 1
            .Header.dwFlags = 0
            .Header.dwLoops = 0
            
            ret = waveOutPrepareHeader(hWaveOut, .Header, Len(.Header))
            
            .Status = ret = MMSYSERR_NOERROR
        End With
         
        If ret Then
            Clear
            err.Raise ERROR_PREPARE_BUFFERS Or ret
            Exit Function
        End If
             
    Next
        
    InitPlayback = True
             
End Function

' // Capture initialization
Public Function InitCapture( _
                ByVal NumOfChannels As Integer, _
                ByVal SamplesPerSec As Long, _
                ByVal BitsPerSample As Integer, _
                ByVal BufferSampleCount As Long, _
                Optional ByVal DeviceID As Long = WAVE_MAPPER, _
                Optional ByVal BuffersCount As Byte = 4) As Boolean
    Dim ret As MMRESULT
    Dim idx As Long
    
    If Not Init Then
        err.Raise Errors.ERROR_OBJECT_FAILED
        Exit Function
    End If
    
    If unavailable Then
        err.Raise Errors.ERROR_UNAVAILABLE
        Exit Function
    End If
    
    If BuffersCount < 1 Then
        err.Raise Errors.INVALID_BUFFERS_COUNT
        Exit Function
    End If
    
    unavailable = True

    With mFormat
        .cbSize = 0
        .wFormatTag = WAVE_FORMAT_PCM
        .wBitsPerSample = BitsPerSample
        .nSamplesPerSec = SamplesPerSec
        .nChannels = NumOfChannels
        .nBlockAlign = .nChannels * .wBitsPerSample \ 8
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
    End With

    mSmpCount = BufferSampleCount - (BufferSampleCount Mod mFormat.nBlockAlign)
    
    ret = waveInOpen(hWaveIn, DeviceID, mFormat, hwnd, 0, CALLBACK_WINDOW)
    
    If ret Then
        err.Raise ERROR_OPEN_DEVICE Or ret
        Exit Function
    End If
    
    bufCount = BuffersCount
    ReDim Buffers(BuffersCount - 1)

    For idx = 0 To BuffersCount - 1
    
        With Buffers(idx)
            ReDim .data(mSmpCount * mFormat.nBlockAlign - 1)
            .Header.lpData = VarPtr(.data(0))
            .Header.dwBufferLength = UBound(.data) + 1
            .Header.dwFlags = 0
            .Header.dwLoops = 0
            
            ret = waveInPrepareHeader(hWaveIn, .Header, Len(.Header))
            
            .Status = ret = MMSYSERR_NOERROR
        End With
         
        If ret Then
            Clear
            err.Raise ERROR_PREPARE_BUFFERS Or ret
            Exit Function
        End If
             
    Next
    
    For idx = 0 To BuffersCount - 1
    
        ret = waveInAddBuffer(hWaveIn, Buffers(idx).Header, Len(Buffers(idx).Header))
        If ret Then
            Clear
            err.Raise ERROR_PREPARE_BUFFERS Or ret
            Exit Function
        End If
        
    Next
    
    InitCapture = True
    
End Function

' // Uninitialize capture/playback
Public Function Uninitialize() As Boolean

    Clear
    Uninitialize = True
    
End Function

' // CopyMemory wrapper
Public Sub CopyData( _
           ByVal pDestination As Long, _
           ByVal pSource As Long, _
           ByVal Length As Long)
    memcpy ByVal pDestination, ByVal pSource, Length
End Sub

' // ------------------------------------------------------------------------------------------------------------

Private Function WndProc( _
                 ByVal hwnd As Long, _
                 ByVal Msg As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Long) As Long
    Dim idx As Long
    Dim hdr As WAVEHDR
    
    If unavailable Then
    
        Select Case Msg
        Case MM_WIM_DATA
            
            memcpy hdr, ByVal lParam, Len(hdr)

            RaiseEvent NewData(hdr.lpData, mSmpCount * mFormat.nBlockAlign)
            
            idx = GetBufferIndex(hdr.lpData)
            
            If idx = -1 Then Exit Function
            
            waveInAddBuffer hWaveIn, Buffers(idx).Header, Len(Buffers(idx).Header)
            
            Exit Function
            
        Case MM_WOM_DONE
            
            If mActive Then
                
                memcpy hdr, ByVal lParam, Len(hdr)
    
                RaiseEvent NewData(hdr.lpData, mSmpCount * mFormat.nBlockAlign)
                
                idx = GetBufferIndex(hdr.lpData)
                
                If idx = -1 Then Exit Function
                
                waveOutWrite hWaveOut, Buffers(idx).Header, Len(Buffers(idx).Header)
                
            End If
            
            Exit Function
            
        End Select
        
    End If
    
    WndProc = DefWindowProc(hwnd, Msg, wParam, lParam)
    
End Function

Private Function CreateAsm() As Boolean
    Dim inIDE   As Boolean
    Dim AsmSize As Long
    Dim ptr     As Long
    Dim isFirst As Boolean

    Debug.Assert MakeTrue(inIDE)
    
    If lpAsm = 0 Then
        If inIDE Then AsmSize = &H2C Else AsmSize = &H20
        hHeap = GetPrevHeap()
        
        If hHeap = 0 Then
            hHeap = HeapCreate(HEAP_CREATE_ENABLE_EXECUTE Or HEAP_NO_SERIALIZE, 0, 0)
            If hHeap = 0 Then err.Raise 7: Exit Function
            If Not SaveCurHeap() Then HeapDestroy hHeap: hHeap = 0: err.Raise 7: Exit Function
            isFirst = True
        End If
        
        lpAsm = HeapAlloc(hHeap, HEAP_NO_SERIALIZE Or HEAP_ZERO_MEMORY, AsmSize)
        
        If lpAsm = 0 Then
            If isFirst Then HeapDestroy hHeap
            hHeap = 0
            err.Raise 7
            Exit Function
        End If
        
    End If
    
    ptr = lpAsm
    
    If inIDE Then
        CreateIDEStub (ptr): ptr = ptr + &HD
    End If
    
    CreateStackConv ptr
    CreateAsm = True
    
End Function

Private Function SaveCurHeap() As Boolean
    Dim i   As Long
    Dim out As String
    
    out = Hex(hHeap)
    For i = Len(out) + 1 To 8: out = "0" & out: Next
    SaveCurHeap = SetEnvironmentVariable(StrPtr(SndClass), StrPtr(out))
    
End Function

Private Function GetPrevHeap() As Long
    Dim out  As String
    
    out = Space(&H8)
    If GetEnvironmentVariable(StrPtr(SndClass), StrPtr(out), LenB(out)) Then GetPrevHeap = Val("&H" & out)
    
End Function

Private Function CreateStackConv( _
                 ByVal ptr As Long) As Boolean
    Dim lpMeth      As Long
    Dim vTable      As Long
    
    GetMem4 ByVal ObjPtr(Me), vTable
    GetMem4 ByVal vTable + WNDPROCINDEX * 4 + &H1C, lpMeth
    
    GetMem4 &H5450C031, ByVal ptr + &H0:    GetMem4 &H488DE409, ByVal ptr + &H4:    GetMem4 &H2474FF04, ByVal ptr + &H8
    GetMem4 &H68FAE018, ByVal ptr + &HC:    GetMem4 &H12345678, ByVal ptr + &H10:   GetMem4 &HFFFFDAE8, ByVal ptr + &H14
    GetMem4 &H10C258FF, ByVal ptr + &H18:   GetMem4 &H0, ByVal ptr + &H1C
    
    GetMem4 ObjPtr(Me), ByVal ptr + &H10                    ' Push Me
    GetMem4 lpMeth - (ptr + &H14) - 5, ByVal ptr + &H14 + 1 ' Call WndProc
    
End Function
 
Private Function CreateIDEStub( _
                 ByVal ptr As Long) As Boolean
    Dim hInstVB6    As Long
    Dim lpEbMode    As Long
    Dim hInstUser32 As Long
    Dim lpDefProc   As Long
    
    hInstVB6 = GetModuleHandle(StrPtr("vba6"))
    If hInstVB6 = 0 Then Exit Function
    hInstUser32 = GetModuleHandle(StrPtr("user32"))
    If hInstUser32 = 0 Then Exit Function
    
    lpEbMode = GetProcAddress(hInstVB6, "EbMode")
    If lpEbMode = 0 Then Exit Function
    lpDefProc = GetProcAddress(hInstUser32, "DefWindowProcW")
    If lpDefProc = 0 Then Exit Function


    GetMem4 &HFFFFFBE8, ByVal ptr + &H0:    GetMem4 &HFC8FEFF, ByVal ptr + &H4
    GetMem4 &H34566B85, ByVal ptr + &H8:    GetMem4 &H12, ByVal ptr + &HC

    GetMem4 lpEbMode - ptr - 5, ByVal ptr + 1 + 0               ' Call EbMode
    GetMem4 lpDefProc - (ptr + &HD), ByVal ptr + &H9            ' JNE  DefWindowProcW
    
    CreateIDEStub = True
    
End Function

Private Function MakeTrue( _
                 ByRef value As Boolean) As Boolean

    value = True
    MakeTrue = True
    
End Function

Private Sub Clear()
    Dim idx As Long
    
    unavailable = False
    
    If hWaveIn Then
        
        waveInReset hWaveIn
        
        For idx = 0 To bufCount - 1
        
            If Buffers(idx).Status Then
                waveInUnprepareHeader hWaveIn, Buffers(idx).Header, Len(Buffers(idx).Header)
            End If
            
        Next
    
        waveInClose hWaveIn
        
    Else
            
        waveOutReset hWaveOut
        
        For idx = 0 To bufCount - 1
        
            If Buffers(idx).Status Then
                waveOutUnprepareHeader hWaveOut, Buffers(idx).Header, Len(Buffers(idx).Header)
            End If
            
        Next
        
        waveOutClose hWaveOut
        
    End If
    
    hWaveIn = 0
    hWaveOut = 0
    paused = False
    mActive = False
    bufCount = 0
    Erase Buffers()
    ZeroMemory mFormat, Len(mFormat)
    
End Sub

Private Function GetBufferIndex( _
                 ByVal ptr As Long) As Long
    Dim idx As Long
    
    For idx = 0 To BuffersCount - 1
    
        If Buffers(idx).Header.lpData = ptr Then
            GetBufferIndex = idx
            Exit Function
        End If
        
    Next
    
    GetBufferIndex = -1
End Function

Private Sub Class_Initialize()
    Dim cls     As WNDCLASSEX
    Dim hUser   As Long
    
    cls.cbSize = Len(cls)
    
    If GetClassInfoEx(App.hInstance, StrPtr(SndClass), cls) = 0 Then
        
        hUser = GetModuleHandle(StrPtr("user32"))
        If hUser = 0 Then Exit Sub
        
        cls.hInstance = App.hInstance
        cls.lpfnwndproc = GetProcAddress(hUser, "DefWindowProcW")
        cls.lpszClassName = StrPtr(SndClass)
        
        If RegisterClassEx(cls) = 0 Then Exit Sub

    End If
    
    If Not CreateAsm() Then Exit Sub
   
    hwnd = CreateWindowEx(0, StrPtr(SndClass), 0, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, App.hInstance, ByVal 0&)
    If hwnd = 0 Then Exit Sub
 
    SetWindowLong hwnd, GWL_WNDPROC, lpAsm
    
    Init = True

End Sub

Private Sub Class_Terminate()
    
    If Not Init Then Exit Sub
    
    Clear
    
    DestroyWindow hwnd
    UnregisterClass StrPtr(SndClass), App.hInstance
    
    If hHeap = 0 Then Exit Sub

    HeapFree hHeap, HEAP_NO_SERIALIZE, ByVal lpAsm
    
End Sub
