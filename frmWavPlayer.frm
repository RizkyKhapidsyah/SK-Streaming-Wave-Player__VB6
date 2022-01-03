VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmWavPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Streaming Wave Player"
   ClientHeight    =   4740
   ClientLeft      =   2685
   ClientTop       =   1770
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWavPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5730
   Begin VB.TextBox txtLength 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2820
      Width           =   2055
   End
   Begin VB.TextBox txtChannels 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2340
      Width           =   2055
   End
   Begin VB.TextBox txtSampleSize 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1980
      Width           =   2055
   End
   Begin VB.TextBox txtSampleFrequency 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1620
      Width           =   2055
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   1260
      Width           =   3615
   End
   Begin ComctlLib.Slider sldMain 
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   3660
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   327682
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   4440
      Top             =   4140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2220
      Top             =   4260
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   4260
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   4260
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open..."
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label lblLength 
      Caption         =   "Length:"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmWavPlayer.frx":000C
      Height          =   1035
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   5775
   End
   Begin VB.Label lblChannels 
      Caption         =   "Mono/Stereo:"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   2400
      Width           =   1515
   End
   Begin VB.Label lblSampleSize 
      Caption         =   "Sample Size:"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label lblSampleFrequency 
      Caption         =   "Sample Frequency:"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   1515
   End
   Begin VB.Label lblFileName 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   795
   End
End
Attribute VB_Name = "frmWavPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cWavPlay As cWavPlayer
Attribute m_cWavPlay.VB_VarHelpID = -1
Private m_bMovingSlider As Boolean

Private Sub pOpenFile(ByVal sFile As String)
Dim bOk As Boolean

On Error Resume Next
   
   bOk = m_cWavPlay.OpenFile(sFile)
   bOk = bOk And (err.Number = 0)
   
   If bOk Then
      If m_cWavPlay.Channels = 1 Then
         txtChannels = "Mono"
      Else
         txtChannels = "Stereo"
      End If
      txtSampleFrequency = format$(m_cWavPlay.SamplesPerSecond / 1000, "#,##0.000") & "kHz"
      txtSampleSize = m_cWavPlay.BitsPerSample & "bits"
      txtFileName = sFile
      txtLength = format$(m_cWavPlay.Length / (m_cWavPlay.SamplesPerSecond * m_cWavPlay.Channels * m_cWavPlay.BitsPerSample / 8), "#,##0.00") & "s"
      
      SetPlayState False
      
      sldMain.Value = 0
      sldMain.Enabled = True
      
      Dim iPos As Long
      For iPos = Len(sFile) To 1 Step -1
         If Mid$(sFile, iPos, 1) = "\" Then
            cdlMain.InitDir = Left$(sFile, iPos - 1)
            Exit For
         End If
      Next iPos
      
   Else
      txtChannels = ""
      txtSampleFrequency = ""
      txtSampleSize = ""
      txtFileName = ""
      txtLength = ""
      SetPlayState False
      sldMain.Value = 0
      sldMain.Enabled = False
      cmdPlay.Enabled = False
   End If

End Sub

Private Sub cmdOpen_Click()
Dim sFile As String

On Error GoTo ErrorHandler

   cdlMain.Filter = "Wave Files (*.wav)|*.wav|All Files (*.*)|*.*"
   cdlMain.ShowOpen
   
   sFile = cdlMain.FileName
   
   pOpenFile sFile
     
   Exit Sub

ErrorHandler:
   If Not err.Number = cdlCancel Then
      MsgBox err.Description, vbInformation
   End If
End Sub

Private Sub SetPlayState(ByVal bState As Boolean)
   cmdPlay.Enabled = Not (bState)
   cmdStop.Enabled = bState
   tmrUpdate.Enabled = bState
End Sub

Private Sub cmdStop_Click()
   m_cWavPlay.StopPlay
   SetPlayState False
End Sub


Private Sub cmdPlay_Click()
   Debug.Print "=============PLAY=========", vbCrLf
   If m_cWavPlay.Play() Then
      SetPlayState True
   End If
End Sub

Private Sub Form_Load()

   Set m_cWavPlay = New cWavPlayer
   m_cWavPlay.Attach Me.hwnd
   m_bMovingSlider = False
   cdlMain.InitDir = App.Path
   sldMain.Min = 0
   sldMain.Max = 100
   pOpenFile App.Path & "\oink.wav"
   
End Sub

Private Sub m_cWavPlay_Complete()
   Debug.Print "COMPLETE"
End Sub

Private Sub sldMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_bMovingSlider = True
End Sub

Private Sub sldMain_MouseUp(Butto5n As Integer, Shift As Integer, X As Single, Y As Single)
   m_cWavPlay.FileSeek sldMain.Value * m_cWavPlay.Length / 100
   m_bMovingSlider = False
End Sub

Private Sub tmrUpdate_Timer()
   If Not (m_bMovingSlider) Then
      If Not (m_cWavPlay.Playing()) Then
         SetPlayState False
      End If
      sldMain.Value = (m_cWavPlay.Position / m_cWavPlay.Length) * 100
   End If
End Sub
