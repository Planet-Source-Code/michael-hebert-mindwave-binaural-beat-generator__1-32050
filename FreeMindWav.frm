VERSION 5.00
Begin VB.Form BinBeat 
   Caption         =   "~ Kyma MindWave ~"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   3480
   Icon            =   "FreeMindWav.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame7 
      Caption         =   "Mask Volume"
      Height          =   615
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Width           =   1815
      Begin VB.HScrollBar scrMaskVol 
         Height          =   255
         LargeChange     =   500
         Left            =   120
         Max             =   0
         Min             =   -5000
         SmallChange     =   50
         TabIndex        =   13
         Top             =   240
         Value           =   -2000
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Masking"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   0
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtMask 
         Alignment       =   2  'Center
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "Inverse Pink"
         ToolTipText     =   "Preselected Mask"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1313
      TabIndex        =   3
      ToolTipText     =   "Click to Pause or Stop"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tone Separation"
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Width           =   1815
      Begin VB.HScrollBar scrBalance 
         Height          =   255
         LargeChange     =   500
         Left            =   120
         Max             =   5000
         SmallChange     =   100
         TabIndex        =   9
         Top             =   240
         Value           =   5000
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frequency"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1335
      Begin VB.TextBox txtSetFreq2 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "165"
         ToolTipText     =   "Edit Frequency 2"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.HScrollBar scrToneVol 
      Height          =   255
      LargeChange     =   500
      Left            =   1680
      Max             =   0
      Min             =   -5000
      SmallChange     =   50
      TabIndex        =   1
      Top             =   360
      Value           =   -2500
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tone Volume"
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtSetFreq1 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "159"
      ToolTipText     =   "Edit Frequency 1"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Click to Generate"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frequency"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      Caption         =   "Function"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   3255
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         ToolTipText     =   "Click to Exit Program"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "BinBeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A program to generate binaural beats for brainwave
'synchronization. The difference between the tones should
'fall within the ranges of 1 to 3 Hz (Delta or sleep state),
'3 to 6 Hz (Theta or deep meditative state) or 6 to 12 Hz
'(Alpha or relaxed alert state). Pink noise masking is used
'to enhance the brainwave entrainment. Volume levels should
'be set fairly low for use with headphones.

'Contrary to popular opinion the brainwave entrainment occurs
'even when used with speakers at higher volume levels.

'Put on your phones, sit back and relax....
'Just don't blame me if you're late for work tomorrow ;-)

Option Explicit

Dim dx As New DirectX8
Dim ds As DirectSound8
Dim dsBuffer(2) As DirectSoundSecondaryBuffer8
Dim SetVolume As Long
Dim increment As Double
Dim inputValue As Double
Dim freq1 As Double
Dim freq2 As Double
Dim fileSize As Integer
Dim fileName1 As String
Dim fileName2 As String

Const Pi = 3.141592654
Const SampleRate = 44100
Const amplitude = 127

Private Sub cmdExit_Click()
    Cleanup
    Unload Me
End Sub

Private Sub Cleanup()
    If Not (dsBuffer(0) Is Nothing) Then dsBuffer(0).Stop
    If Not (dsBuffer(1) Is Nothing) Then dsBuffer(1).Stop
    If Not (dsBuffer(2) Is Nothing) Then dsBuffer(2).Stop
    Set dsBuffer(0) = Nothing
    Set dsBuffer(1) = Nothing
    Set dsBuffer(2) = Nothing
    Set ds = Nothing
    Set dx = Nothing
End Sub

Private Sub Form_Load()
    Me.Show
    On Local Error Resume Next
    Set ds = dx.DirectSoundCreate("")
    If Err.Number <> 0 Then
        MsgBox "Unable to start DirectSound"
        End
    End If
    ds.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
    cmdGenerate_Click
End Sub

Private Sub txtSetFreq1_Change()
    freq1 = Val(txtSetFreq1.Text)
End Sub

Private Sub txtSetFreq2_Change()
    freq2 = Val(txtSetFreq2.Text)
End Sub

Private Sub scrBalance_Scroll()
    panVal = scrBalance.Value
    dsBuffer(0).SetPan panVal
    dsBuffer(1).SetPan panVal * -1
End Sub

Private Sub scrBalance_Change()
    panVal = scrBalance.Value
    dsBuffer(0).SetPan panVal
    dsBuffer(1).SetPan panVal * -1
End Sub

Private Sub scrToneVol_Scroll()
    dsBuffer(0).SetVolume scrToneVol.Value
    dsBuffer(1).SetVolume scrToneVol.Value
End Sub

Private Sub scrToneVol_Change()
    dsBuffer(0).SetVolume scrToneVol.Value
    dsBuffer(1).SetVolume scrToneVol.Value
End Sub

Private Sub scrMaskVol_Scroll()
    If Not (dsBuffer(2) Is Nothing) Then
        dsBuffer(2).SetVolume scrMaskVol.Value
    End If
End Sub

Private Sub scrMaskVol_Change()
    If Not (dsBuffer(2) Is Nothing) Then
        dsBuffer(2).SetVolume scrMaskVol.Value
    End If
End Sub

Private Sub cmdGenerate_Click()
    freq1 = Val(txtSetFreq1.Text)
    If freq1 < 10 Or freq1 > 600 Then
        MsgBox "Frequency must be within 10 to 600 Hz range."
        Exit Sub
    End If
    SineWave1
    freq2 = Val(txtSetFreq2.Text)
    If freq2 < 10 Or freq2 > 600 Then
        MsgBox "Frequency must be within 10 to 600 Hz range."
        Exit Sub
    End If
    SineWave2
    Play
End Sub

Private Sub Play()
    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC Or DSBCAPS_GLOBALFOCUS
    fileName1 = App.Path & "\freq1.wav"
    Set dsBuffer(0) = ds.CreateSoundBufferFromFile(fileName1, bufferDesc)
    Close #1
    fileName2 = App.Path & "\freq2.wav"
    Set dsBuffer(1) = ds.CreateSoundBufferFromFile(fileName2, bufferDesc)
    Close #1
    Set dsBuffer(2) = ds.CreateSoundBufferFromFile(App.Path & "\pink_inv.wav", bufferDesc)
    dsBuffer(0).SetVolume scrToneVol.Value
    dsBuffer(0).Play DSBPLAY_LOOPING
    dsBuffer(1).SetVolume scrToneVol.Value
    dsBuffer(1).Play DSBPLAY_LOOPING
    If Not (dsBuffer(2) Is Nothing) Then
        dsBuffer(2).SetVolume scrMaskVol.Value
        dsBuffer(2).Play DSBPLAY_LOOPING
    End If
End Sub

Private Sub cmdStop_Click()
    dsBuffer(0).Stop
    dsBuffer(1).Stop
    If Not (dsBuffer(2) Is Nothing) Then
        dsBuffer(2).Stop
    End If
End Sub

Private Sub MakeFile1()
    fileName1 = App.Path & "\freq1.wav"
    Kill fileName1
    Open fileName1 For Binary Access Write As #1
        Put #1, 1, "RIFF"
        Put #1, 5, CInt(0)          'Will write length later
        Put #1, 9, "WAVEfmt "
        Put #1, 17, CLng(16)        'Lenth of format data
        Put #1, 21, CInt(1)         'Wave type PCM
        Put #1, 23, CInt(1)         '1 channel
        Put #1, 25, CLng(44100)     '44.1 hHz SampleRate
        Put #1, 29, CLng(88200)
        Put #1, 33, CInt(2)
        Put #1, 35, CInt(16)
        Put #1, 37, "data"
        Put #1, 41, CInt(0)
End Sub

Private Sub MakeFile2()
    fileName2 = App.Path & "\freq2.wav"
    Kill fileName2
    Open fileName2 For Binary Access Write As #1
        Put #1, 1, "RIFF"
        Put #1, 5, CInt(0)          'Will write length later
        Put #1, 9, "WAVEfmt "
        Put #1, 17, CLng(16)        'Lenth of format data
        Put #1, 21, CInt(1)         'Wave type PCM
        Put #1, 23, CInt(1)         '1 channel
        Put #1, 25, CLng(44100)     '44.1 hHz SampleRate
        Put #1, 29, CLng(88200)
        Put #1, 33, CInt(2)
        Put #1, 35, CInt(16)
        Put #1, 37, "data"
        Put #1, 41, CInt(0)
End Sub

Private Sub SineWave1()
    MakeFile1
    Dim sample As Integer
    Dim increment As Double
    Dim bufferbyte As Integer
        bufferbyte = 45
        increment = Pi / (SampleRate / freq1)
        For inputValue = 0 To (2 * Pi) Step increment
            sample = Int(amplitude * Sin(inputValue))
            Put #1, bufferbyte, sample
            bufferbyte = bufferbyte + 1
        Next inputValue
    CloseFile
End Sub

Private Sub SineWave2()
    MakeFile2
    Dim sample As Integer
    Dim increment As Double
    Dim bufferbyte As Integer
        bufferbyte = 45
        increment = Pi / (SampleRate / freq2)
        For inputValue = 0 To (2 * Pi) Step increment
            sample = Int(amplitude * Sin(inputValue))
            Put #1, bufferbyte, sample
            bufferbyte = bufferbyte + 1
        Next inputValue
    CloseFile
End Sub

Private Sub CloseFile()
    fileSize = LOF(1)
    Put #1, 5, CLng(fileSize - 8)
    Put #1, 41, CLng(fileSize - 44)
    Close #1
End Sub
