VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWavPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_DESTROY = &H2&

Private Const CALLBACK_WINDOW = &H10000
Private Const WAVE_MAPPER = -1&
Private Const MMIO_READ = &H0
Private Const MMIO_FINDCHUNK = &H10
Private Const MMIO_FINDRIFF = &H20
Private Const MM_WOM_DONE = &H3BD
Private Const MMSYSERR_NOERROR = 0
Private Const SEEK_CUR = 1
Private Const SEEK_END = 2
Private Const SEEK_SET = 0
Private Const TIME_BYTES = &H4
Private Const WHDR_DONE = &H1

' Wave formats:
Public Enum EWaveFormat
   WAVE_FORMAT_UNKNOWN = &H0        ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_PCM = &H1
   WAVE_FORMAT_ADPCM = &H2          ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_IEEE_FLOAT = &H3     ' /*  Microsoft Corporation  *' /
                                        ' /*  IEEE754: range (+1, -1]  *' /
                                        ' /*  32-bit' /64-bit format as defined by *' /
                                        ' /*  MSVC++ float' /double type *' /
   WAVE_FORMAT_IBM_CVSD = &H5       ' /*  IBM Corporation  *' /
   WAVE_FORMAT_ALAW = &H6           ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_MULAW = &H7          ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_OKI_ADPCM = &H10     ' /*  OKI  *' /
   WAVE_FORMAT_DVI_ADPCM = &H11     ' /*  Intel Corporation  *' /
   WAVE_FORMAT_MEDIASPACE_ADPCM = &H12      ' /*  Videologic  *' /
   WAVE_FORMAT_SIERRA_ADPCM = &H13          ' /*  Sierra Semiconductor Corp  *' /
   WAVE_FORMAT_G723_ADPCM = &H14    ' /*  Antex Electronics Corporation  *' /
   WAVE_FORMAT_DIGISTD = &H15       ' /*  DSP Solutions, Inc.  *' /
   WAVE_FORMAT_DIGIFIX = &H16       ' /*  DSP Solutions, Inc.  *' /
   WAVE_FORMAT_DIALOGIC_OKI_ADPCM = &H17    ' /*  Dialogic Corporation  *' /
   WAVE_FORMAT_MEDIAVISION_ADPCM = &H18     ' /*  Media Vision, Inc. *' /
   WAVE_FORMAT_YAMAHA_ADPCM = &H20          ' /*  Yamaha Corporation of America  *' /
   WAVE_FORMAT_SONARC = &H21        ' /*  Speech Compression  *' /
   WAVE_FORMAT_DSPGROUP_TRUESPEECH = &H22           ' /*  DSP Group, Inc  *' /
   WAVE_FORMAT_ECHOSC1 = &H23       ' /*  Echo Speech Corporation  *' /
   WAVE_FORMAT_AUDIOFILE_AF36 = &H24        ' /*    *' /
   WAVE_FORMAT_APTX = &H25          ' /*  Audio Processing Technology  *' /
   WAVE_FORMAT_AUDIOFILE_AF10 = &H26        ' /*    *' /
   WAVE_FORMAT_DOLBY_AC2 = &H30     ' /*  Dolby Laboratories  *' /
   WAVE_FORMAT_GSM610 = &H31        ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_MSNAUDIO = &H32      ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_ANTEX_ADPCME = &H33          ' /*  Antex Electronics Corporation  *' /
   WAVE_FORMAT_CONTROL_RES_VQLPC = &H34     ' /*  Control Resources Limited  *' /
   WAVE_FORMAT_DIGIREAL = &H35      ' /*  DSP Solutions, Inc.  *' /
   WAVE_FORMAT_DIGIADPCM = &H36     ' /*  DSP Solutions, Inc.  *' /
   WAVE_FORMAT_CONTROL_RES_CR10 = &H37      ' /*  Control Resources Limited  *' /
   WAVE_FORMAT_NMS_VBXADPCM = &H38          ' /*  Natural MicroSystems  *' /
   WAVE_FORMAT_CS_IMAADPCM = &H39   ' /* Crystal Semiconductor IMA ADPCM *' /
   WAVE_FORMAT_ECHOSC3 = &H3A       ' /* Echo Speech Corporation *' /
   WAVE_FORMAT_ROCKWELL_ADPCM = &H3B        ' /* Rockwell International *' /
   WAVE_FORMAT_ROCKWELL_DIGITALK = &H3C     ' /* Rockwell International *' /
   WAVE_FORMAT_XEBEC = &H3D         ' /* Xebec Multimedia Solutions Limited *' /
   WAVE_FORMAT_G721_ADPCM = &H40    ' /*  Antex Electronics Corporation  *' /
   WAVE_FORMAT_G728_CELP = &H41     ' /*  Antex Electronics Corporation  *' /
   WAVE_FORMAT_MPEG = &H50          ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_MPEGLAYER3 = &H55    ' /*  ISO' /MPEG Layer3 Format Tag *' /
   WAVE_FORMAT_CIRRUS = &H60        ' /*  Cirrus Logic  *' /
   WAVE_FORMAT_ESPCM = &H61         ' /*  ESS Technology  *' /
   WAVE_FORMAT_VOXWARE = &H62       ' /*  Voxware Inc  *' /
   WAVEFORMAT_CANOPUS_ATRAC = &H63          ' /*  Canopus, co., Ltd.  *' /
   WAVE_FORMAT_G726_ADPCM = &H64    ' /*  APICOM  *' /
   WAVE_FORMAT_G722_ADPCM = &H65    ' /*  APICOM      *' /
   WAVE_FORMAT_DSAT = &H66          ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_DSAT_DISPLAY = &H67          ' /*  Microsoft Corporation  *' /
   WAVE_FORMAT_SOFTSOUND = &H80     ' /*  Softsound, Ltd.      *' /
   WAVE_FORMAT_RHETOREX_ADPCM = &H100       ' /*  Rhetorex Inc  *' /
   WAVE_FORMAT_CREATIVE_ADPCM = &H200       ' /*  Creative Labs, Inc  *' /
   WAVE_FORMAT_CREATIVE_FASTSPEECH8 = &H202         ' /*  Creative Labs, Inc  *' /
   WAVE_FORMAT_CREATIVE_FASTSPEECH10 = &H203        ' /*  Creative Labs, Inc  *' /
   WAVE_FORMAT_QUARTERDECK = &H220  ' /*  Quarterdeck Corporation  *' /
   WAVE_FORMAT_FM_TOWNS_SND = &H300         ' /*  Fujitsu Corp.  *' /
   WAVE_FORMAT_BTV_DIGITAL = &H400          ' /*  Brooktree Corporation  *' /
   WAVE_FORMAT_OLIGSM = &H1000      ' /*  Ing C. Olivetti & C., S.p.A.  *' /
   WAVE_FORMAT_OLIADPCM = &H1001    ' /*  Ing C. Olivetti & C., S.p.A.  *' /
   WAVE_FORMAT_OLICELP = &H1002     ' /*  Ing C. Olivetti & C., S.p.A.  *' /
   WAVE_FORMAT_OLISBC = &H1003      ' /*  Ing C. Olivetti & C., S.p.A.  *' /
   WAVE_FORMAT_OLIOPR = &H1004      ' /*  Ing C. Olivetti & C., S.p.A.  *' /
   WAVE_FORMAT_LH_CODEC = &H1100    ' /*  Lernout & Hauspie  *' /
   WAVE_FORMAT_NORRIS = &H1400      ' /*  Norris Communications, Inc.  *' /

' /' /
' /' /  the WAVE_FORMAT_DEVELOPMENT format tag can be used during the
' /' /  development phase of a new wave format.  Before shipping, you MUST
' /' /  acquire an official format tag from Microsoft.
' /' /
   WAVE_FORMAT_DEVELOPMENT = &HFFFF
End Enum

Private Type mmioinfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type

Private Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
End Type

Private Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
End Type

Private Type WAVEFORMATEX
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Private Type MMCKINFO
   ckid As Long
   ckSize As Long
   fccType As Long
   dwDataOffset As Long
   dwFlags As Long
End Type

Private Type MMTIME
   wType As Long
   u As Long
   X As Long
End Type

Private Declare Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Private Declare Function waveOutOpen Lib "winmm.dll" (hWaveOut As Long, ByVal uDeviceID As Long, ByVal format As String, ByVal dwCallback As Long, ByRef fPlaying As Boolean, ByVal dwFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Private Declare Function mmioReadString Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Private Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal cb As Long)
Private Declare Sub CopyMemoryFromString Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal source As String, ByVal cb As Long)
Private Declare Function PostWavMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef hdr As WAVEHDR) As Long
Private Const GMEM_FIXED = &H0

Implements ISubclass

Private Const NUM_BUFFERS = 5
Private Const BUFFER_SECONDS = 0.1
Private Const cErrBase = 29670

Private m_hWnd As Long

Private m_bPlaying As Boolean

Private m_fBufferLen As Single

Private m_hMmioIn As Long      ' file handle

Private m_lStartPos As Long    ' sample where we started playback from
Private m_lDataOffset As Long  ' start of audio data in wave file
Private m_lAudioLength As Long  ' number of bytes in audio data

Private m_lPtrFormat As Long      ' pointer to wave format
Private m_tFormat As WAVEFORMATEX    ' waveformat structure

Private m_hMem(1 To NUM_BUFFERS) As Long      ' memory handles
Private m_lPtrMem(1 To NUM_BUFFERS) As Long      ' memory pointers
Private m_tHdr(1 To NUM_BUFFERS) As WAVEHDR    ' wave headers
Private m_lBufferSize As Long       ' size of output buffers

Private m_hWaveOut As Long        ' waveout handle

Public Event Complete()

Public Property Get BufferLength() As Single
   BufferLength = m_fBufferLen
End Property
Public Property Let BufferLength(ByVal f As Single)
   If f > 0.001 And f < 10 Then
      m_fBufferLen = f
   Else
      pInternalErrorHandler 10
   End If
End Property

Public Property Get WaveFormatTag() As EWaveFormat
   If m_hMmioIn Then
      WaveFormatTag = m_tFormat.wFormatTag
   Else
      pInternalErrorHandler 6
   End If
End Property
Public Property Get SamplesPerSecond() As Long
   If m_hMmioIn Then
      SamplesPerSecond = m_tFormat.nSamplesPerSec
   Else
      pInternalErrorHandler 6
   End If
End Property
Public Property Get Channels() As Long
   If m_hMmioIn Then
      Channels = m_tFormat.nChannels
   Else
      pInternalErrorHandler 6
   End If
End Property
Public Property Get BitsPerSample() As Long
   If m_hMmioIn Then
      BitsPerSample = m_tFormat.wBitsPerSample
   Else
      pInternalErrorHandler 6
   End If
End Property
Public Sub CloseFile()
   If m_hMmioIn Then
      StopPlay
      mmioClose m_hMmioIn, 0
      m_hMmioIn = 0
   End If
End Sub
Public Function OpenFile(ByVal sSoundFile As String) As Boolean
Dim lR As Long
Dim mmckinfoParentIn As MMCKINFO
Dim mmckinfoSubchunkIn As MMCKINFO
Dim mmioinf As mmioinfo
Dim sFormat As String
Dim iBuffer As Long
Dim bFailed As Boolean
Dim lRem As Long
    
   ' close previously open file (if any)
   CloseFile
   
   If (sSoundFile = "") Then
      Exit Function
   End If
        
   ' Open the input file
   m_hMmioIn = mmioOpen(sSoundFile, mmioinf, MMIO_READ)
   If (m_hMmioIn = 0) Then
      pInternalErrorHandler 2
      Exit Function
   End If

   ' Check if this is a wave file
   mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
   lR = mmioDescendParent(m_hMmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
   If Not (lR = MMSYSERR_NOERROR) Then
      CloseFile
      pInternalErrorHandler 3
      Exit Function
   End If

   ' Get format info
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
   lR = mmioDescend(m_hMmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (lR <> MMSYSERR_NOERROR) Then
      CloseFile
      pInternalErrorHandler 4
      Exit Function
   End If
   
   sFormat = String$(50, 0)
   lR = mmioReadString(m_hMmioIn, sFormat, mmckinfoSubchunkIn.ckSize)
   If (lR = -1) Then
      CloseFile
      pInternalErrorHandler 5
      Exit Function
   End If
   lR = mmioAscend(m_hMmioIn, mmckinfoSubchunkIn, 0)
   CopyMemoryFromString m_tFormat, sFormat, Len(m_tFormat)
    
   ' Find the data subchunk
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
   lR = mmioDescend(m_hMmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If Not (lR = MMSYSERR_NOERROR) Then
      CloseFile
      pInternalErrorHandler 6
      Exit Function
   End If
   m_lDataOffset = mmioSeek(m_hMmioIn, 0, SEEK_CUR)
    
   ' Get the length of the audio
   m_lAudioLength = mmckinfoSubchunkIn.ckSize
    
   ' Allocate audio buffers - we aim to create a buffer long
   ' enough to cope with BUFFER_SECONDS of sound.
   m_lBufferSize = m_tFormat.nSamplesPerSec * m_fBufferLen
   If m_tFormat.wBitsPerSample > 8 Then
      m_lBufferSize = m_lBufferSize * (m_tFormat.wBitsPerSample \ 2)
   End If
   If m_tFormat.nChannels > 0 Then
      m_lBufferSize = m_lBufferSize * m_tFormat.nChannels
   End If
   If m_lBufferSize = 0 Then
      ' assume stereo 16bit 44.1k for sake of argument:
      m_lBufferSize = 44100 * 2 * 2 * BUFFER_SECONDS
   End If
   ' Align to block alignment boundary:
   If (m_lBufferSize Mod m_tFormat.nBlockAlign) > 0 Then
      lRem = (m_lBufferSize \ m_tFormat.nBlockAlign) + 1
      m_lBufferSize = lRem * m_tFormat.nBlockAlign
   End If
      
   pClearBuffers
   
   For iBuffer = 1 To NUM_BUFFERS
      m_hMem(iBuffer) = GlobalAlloc(GMEM_FIXED, m_lBufferSize)
      If m_hMem(iBuffer) = 0 Then
         bFailed = True
         Exit For
      Else
         m_lPtrMem(iBuffer) = GlobalLock(m_hMem(iBuffer))
         If m_lPtrMem(iBuffer) = 0 Then
            bFailed = True
            Exit For
         End If
      End If
    Next iBuffer
    
    If bFailed Then
      ' Clear up buffers:
      pClearBuffers
      ' Close File:
      CloseFile
      ' Return failure:
      pInternalErrorHandler 7
   Else
      OpenFile = True
   End If
    
End Function
Private Sub pClearBuffers()
Dim iBuffer As Long
   For iBuffer = 1 To NUM_BUFFERS
      If Not m_lPtrMem(iBuffer) = 0 Then
         GlobalUnlock m_lPtrMem(iBuffer)
         m_lPtrMem(iBuffer) = 0
      End If
      If Not m_hMem(iBuffer) = 0 Then
         GlobalFree m_hMem(iBuffer)
         m_hMem(iBuffer) = 0
      End If
   Next iBuffer
End Sub

Public Sub Attach(ByVal hwnd As Long)
   m_hWnd = hwnd
   AttachMessage Me, m_hWnd, MM_WOM_DONE
   AttachMessage Me, m_hWnd, WM_DESTROY
End Sub
Public Sub Detach()
   If m_hWnd Then
      CloseFile
      pClearBuffers
      DetachMessage Me, m_hWnd, MM_WOM_DONE
      DetachMessage Me, m_hWnd, WM_DESTROY
      m_hWnd = 0
   End If
End Sub
Private Sub pErrorHandler(ByVal lR As Long)
Dim sMsg As String
Dim iPos As Long
   sMsg = String$(260, 0)
   waveOutGetErrorText lR, sMsg, Len(sMsg)
   iPos = InStr(sMsg, vbNullChar)
   If Not iPos = 0 Then
      sMsg = Left$(sMsg, iPos - 1)
   End If
   err.Raise cErrBase + lR + 10, App.EXEName & ".cWavePlayer", sMsg
End Sub
Private Sub pInternalErrorHandler(ByVal lR As Long)
Dim sMsg As String
   Select Case lR
   Case 1
      sMsg = "Class not ready; call attach method first."
   Case 2
      sMsg = "Unable to open file."
   Case 3
      sMsg = "Not a Wave file."
   Case 4
      sMsg = "Unable to retrieve format chunk"
   Case 5
      sMsg = "Error reading format"
   Case 6
      sMsg = "No Wave File Open"
   Case 7
      sMsg = "Insufficient memory"
   Case 8
      sMsg = "Position out of range"
   Case 9
      sMsg = "No wave file playing"
   Case 10
      sMsg = "Buffer time out of range, choose a number of second between 0.001 and 10"
   End Select
   err.Raise cErrBase + lR, App.EXEName & ".cWavePlayer", sMsg
End Sub

Public Function Play() As Boolean
Dim lR As Long
Dim sFormatBuffer As String
Dim iBuffer As Long

   If m_hWnd = 0 Then
      pInternalErrorHandler 1
      Exit Function
   End If
   If m_hMmioIn = 0 Then
      pInternalErrorHandler 6
      Exit Function
   End If

   If (m_bPlaying) Then
      Play = True
      Exit Function
   End If
       
   sFormatBuffer = String$(50, 0)
   CopyMemory ByVal sFormatBuffer, m_tFormat, LenB(m_tFormat)
   lR = waveOutOpen(m_hWaveOut, WAVE_MAPPER, sFormatBuffer, m_hWnd, True, CALLBACK_WINDOW)
    
   If Not (lR = MMSYSERR_NOERROR) Then
      pErrorHandler lR
      Play = False
      Exit Function
   End If

   For iBuffer = 1 To NUM_BUFFERS
      With m_tHdr(iBuffer)
         .lpData = m_lPtrMem(iBuffer)
         .dwBufferLength = m_lBufferSize
         .dwUser = iBuffer
         .dwFlags = 0
         .dwLoops = 0
      End With
      lR = waveOutPrepareHeader(m_hWaveOut, m_tHdr(iBuffer), LenB(m_tHdr(iBuffer)))
      If Not (lR = MMSYSERR_NOERROR) Then
         pErrorHandler lR
      End If
   Next iBuffer

   m_bPlaying = True
   Play = True

   m_lStartPos = mmioSeek(m_hMmioIn, 0, SEEK_CUR) - m_lDataOffset
    
   ' Start playing by posting callback functions to read into
   ' the five buffers & play:
   For iBuffer = 1 To NUM_BUFFERS
      PostWavMessage m_hWnd, MM_WOM_DONE, 0, m_tHdr(iBuffer)
   Next
    
End Function

Public Sub StopPlay()
   If m_bPlaying Then
      m_bPlaying = False
      FileSeek Position
      waveOutReset m_hWaveOut
   End If
End Sub

Public Property Get Length() As Long
   ' Returns the length in bytes:
   Length = m_lAudioLength
End Property

Public Function FileSeek(ByVal Position As Long) As Boolean
Dim lBytePos As Long
Dim lR As Long
    
   FileSeek = False
   lBytePos = Position
   If (m_hMmioIn = 0) Then
      pInternalErrorHandler 6
      Exit Function
   Else
      If (lBytePos < 0) Or (lBytePos >= m_lAudioLength) Then
         pInternalErrorHandler 8
         Exit Function
      End If
   End If
   
   ' Ensure position is on a byte boundary:
   If Not (lBytePos Mod m_tFormat.nBlockAlign) = 0 Then
      lBytePos = ((lBytePos \ m_tFormat.nBlockAlign) + 1) * m_tFormat.nBlockAlign
   End If
   lR = mmioSeek(m_hMmioIn, lBytePos + m_lDataOffset, SEEK_SET)
   m_lStartPos = lR
   FileSeek = True
   
End Function

Public Sub Pause(ByVal bState As Boolean)
   If m_hMmioIn = 0 Then
      pInternalErrorHandler 6
      Exit Sub
   End If
   If m_hWaveOut = 0 Then
      pInternalErrorHandler 9
   End If
   If bState Then
      waveOutPause m_hWaveOut
   Else
      waveOutRestart m_hWaveOut
   End If
End Sub

Public Property Get Position() As Long
Dim tMMT As MMTIME
Dim lR As Long

   tMMT.wType = TIME_BYTES
   lR = waveOutGetPosition(m_hWaveOut, tMMT, LenB(tMMT))
   If (lR = MMSYSERR_NOERROR) Then
      Position = (m_lStartPos + tMMT.u)
   'Else
   '   Position = (mmioSeek(m_hMmioIn, 0, SEEK_CUR) - m_lDataOffset + m_lBufferSize * NUM_BUFFERS) \ m_tFormat.nBlockAlign
   End If
   
End Property

Public Property Get Playing() As Boolean
Dim tMMT As MMTIME
Dim lR As Long
    
   If Not (m_hMmioIn = 0 Or m_hWaveOut = 0) Then
      tMMT.wType = TIME_BYTES
      lR = waveOutGetPosition(m_hWaveOut, tMMT, LenB(tMMT))
      If (lR = MMSYSERR_NOERROR) Then
         Playing = True
      Else
         Playing = False
      End If
   End If
End Property


Private Sub Class_Initialize()
   m_fBufferLen = BUFFER_SECONDS
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   Detach
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
   ' Windows processes messages first:
   ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Static dataRemaining As Long
Dim tWavHdr As WAVEHDR
Dim lR As Long
Dim iBuffer As Long
    
   Select Case iMsg
   Case MM_WOM_DONE
   
      ' Get the WAVEHDR structure for this call:
      CopyMemory tWavHdr, ByVal lParam, LenB(tWavHdr)
      
      'Debug.Print tWavHdr.lpData, tWavHdr.dwBufferLength, tWavHdr.dwUser
      
      ' Send data to the output on this buffer (if we have data remaining
      ' and we are playing.  If we reach the end of the wave file, then signal
      ' that we are no longer playing and prepare to clear up the buffers:
      If (m_bPlaying) Then
         dataRemaining = (m_lDataOffset + m_lAudioLength - mmioSeek(m_hMmioIn, 0, SEEK_CUR))
         If (m_lBufferSize < dataRemaining) Then
            ' Get m_lBufferSize bytes from the WAV file into the memory
            ' buffer pointed to by tWavHdr.lpData:
            lR = mmioRead(m_hMmioIn, tWavHdr.lpData, m_lBufferSize)
         Else
            ' Get the remainder (dataRemaining) bytes from the WAV
            ' file into the memory buffer pointed to by tWavHdr.lpData:
            lR = mmioRead(m_hMmioIn, tWavHdr.lpData, dataRemaining)
            m_bPlaying = False
         End If
         ' Write the amount of data we just read into the
         ' memory buffer to the output sound device output buffer.  The
         ' waveOutWrite call will return immediately whilst the output
         ' device works on the data we just provided:
         tWavHdr.dwBufferLength = lR
         lR = waveOutWrite(m_hWaveOut, tWavHdr, LenB(tWavHdr))
      End If
      
      ' Ensure we close all buffers
      If Not (m_bPlaying) Then
         'Debug.Print "Closing buffer ", tWavHdr.dwUser
         
         ' Once we have finished with the buffer we can close the buffer:
         waveOutUnprepareHeader m_hWaveOut, m_tHdr(tWavHdr.dwUser), Len(m_tHdr(tWavHdr.dwUser))
         
         ' This is a bit of a kludge, really we should only call waveOutClose once
         ' all the buffers have been cleared, however we can attempt to call
         ' waveOutClose and have it return error 33 whilst there are still
         ' outstanding buffers being played:
         lR = waveOutClose(m_hWaveOut)
         If (lR = MMSYSERR_NOERROR) Then
            RaiseEvent Complete
         End If
         
      End If
      
   Case WM_DESTROY
      ' The app is closing but this class is still attached;
      ' we should try and clear up as a courtesy to the developer:
      On Error Resume Next
      Detach
   End Select
   
End Function
