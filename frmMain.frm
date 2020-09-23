VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Audio Tone Generator"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write Wave File..."
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   5085
      Width           =   1860
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wave Type"
      Height          =   1995
      Left            =   8145
      TabIndex        =   25
      Top             =   1080
      Width           =   1590
      Begin VB.OptionButton optSignal 
         Caption         =   "Sine Wave"
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   28
         ToolTipText     =   "Default"
         Top             =   315
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optSignal 
         Caption         =   "Cosine Wave"
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   27
         Top             =   720
         Width           =   1320
      End
      Begin VB.OptionButton optSignal 
         Caption         =   "Combined Wave"
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   26
         Top             =   1125
         Width           =   1320
      End
   End
   Begin VB.VScrollBar scrVol 
      Height          =   3885
      Left            =   7875
      Max             =   0
      Min             =   100
      TabIndex        =   22
      Top             =   90
      Value           =   100
      Width           =   285
   End
   Begin VB.PictureBox pbxScale 
      Height          =   240
      Left            =   90
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   20
      ToolTipText     =   "Window size of 1 second"
      Top             =   4500
      Width           =   7755
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   7155
      Top             =   4005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkCustom 
      Caption         =   "Create from custom drawn time slice"
      Enabled         =   0   'False
      Height          =   780
      Left            =   8190
      TabIndex        =   19
      ToolTipText     =   "Trace bitmap to enable"
      Top             =   3150
      Width           =   1545
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Bitmap"
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Top             =   4725
      Width           =   1860
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Bitmap"
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   4365
      Width           =   1860
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview Waveform"
      Default         =   -1  'True
      Height          =   375
      Left            =   6030
      TabIndex        =   11
      Top             =   5085
      Width           =   1860
   End
   Begin VB.TextBox txtTime 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Text            =   "2"
      Top             =   4770
      Width           =   555
   End
   Begin VB.TextBox txtFreq 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1215
      TabIndex        =   8
      Text            =   "500"
      Top             =   4770
      Width           =   825
   End
   Begin VB.OptionButton optMono 
      Caption         =   "Mono"
      Height          =   285
      Left            =   5985
      TabIndex        =   6
      Top             =   4770
      Width           =   915
   End
   Begin VB.OptionButton optStereo 
      Caption         =   "Stereo"
      Height          =   285
      Left            =   6930
      TabIndex        =   5
      Top             =   4770
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.ComboBox cboBitsPer 
      Height          =   315
      ItemData        =   "frmMain.frx":044A
      Left            =   3285
      List            =   "frmMain.frx":0454
      TabIndex        =   2
      Text            =   "16"
      Top             =   4770
      Width           =   690
   End
   Begin VB.ComboBox cboSampleRate 
      Height          =   315
      ItemData        =   "frmMain.frx":045F
      Left            =   1620
      List            =   "frmMain.frx":047E
      TabIndex        =   1
      Text            =   "44100"
      Top             =   5130
      Width           =   960
   End
   Begin VB.PictureBox pbxWave 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      ForeColor       =   &H0000FF00&
      Height          =   3895
      Left            =   90
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      ToolTipText     =   "Draw on picturebox to create a custom waveform"
      Top             =   90
      Width           =   7735
      Begin VB.Line linCenter 
         BorderColor     =   &H00008000&
         X1              =   0
         X2              =   510
         Y1              =   128
         Y2              =   128
      End
   End
   Begin VB.CommandButton cmdTrace 
      Caption         =   "Trace Bitmap"
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   4005
      Width           =   1860
   End
   Begin VB.Label lblVol 
      Alignment       =   1  'Right Justify
      Caption         =   "100%"
      Height          =   240
      Left            =   8640
      TabIndex        =   24
      Top             =   315
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Master Volume"
      Height          =   195
      Left            =   8640
      TabIndex        =   23
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label lblPos 
      Caption         =   "0 mSec"
      Height          =   330
      Left            =   8190
      TabIndex        =   21
      Top             =   675
      Width           =   1545
   End
   Begin VB.Label lblSize 
      Caption         =   "--"
      Height          =   285
      Left            =   3825
      TabIndex        =   18
      ToolTipText     =   "Click Preview to calculate"
      Top             =   5130
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Approx. File Size"
      Height          =   285
      Left            =   2610
      TabIndex        =   17
      Top             =   5130
      Width           =   1185
   End
   Begin VB.Label lblTimeSlice 
      Alignment       =   2  'Center
      Caption         =   "11.61 mSec"
      Height          =   240
      Left            =   3465
      TabIndex        =   13
      ToolTipText     =   "Number of milliseconds visable"
      Top             =   4140
      Width           =   1185
   End
   Begin VB.Line Line4 
      X1              =   7830
      X2              =   4680
      Y1              =   4275
      Y2              =   4275
   End
   Begin VB.Line Line3 
      X1              =   90
      X2              =   3420
      Y1              =   4275
      Y2              =   4275
   End
   Begin VB.Line Line2 
      X1              =   7830
      X2              =   7830
      Y1              =   4095
      Y2              =   4605
   End
   Begin VB.Line Line1 
      X1              =   90
      X2              =   90
      Y1              =   4050
      Y2              =   4545
   End
   Begin VB.Label Label4 
      Caption         =   "Duration (Seconds)"
      Height          =   285
      Left            =   4005
      TabIndex        =   10
      Top             =   4770
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "Frequency (Hz)"
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Samples Per Second"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   5130
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Bits Per Sample"
      Height          =   285
      Left            =   2070
      TabIndex        =   3
      Top             =   4770
      Width           =   1185
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This code is designed to allow the user to create wave samples through generating
'sine, cosine, and combination waveforms. It also allows drawing a custom waveform
'to produce a unique sound. I have included a few extras such as writing to both channels.
'This is to show how writing left and right can be implemented.

Dim WaveH As WAVEHEADER
Dim WaveF As FORMATHEADER
Dim WaveD As DATAHEADER

Dim Custom(0 To 511) As Integer
Dim WaveType As Integer

'use as a boolean
Dim DrawOn As Byte

Private Sub cboBitsPer_Change()
    WaveF.BitsPerSample = Val(cboBitsPer.Text)
End Sub

Private Sub cboBitsPer_Click()
    cboBitsPer_Change
End Sub

Private Sub cboBitsPer_Scroll()
    cboBitsPer_Change
End Sub

Private Sub cboSampleRate_Change()
    
    If cboSampleRate.Text = "" Then
        cboSampleRate.Text = "44100"
    ElseIf cboSampleRate.Text = "0" Then
        cboSampleRate.Text = "44100"
    End If
    
    Ratio = 512 / Val(cboSampleRate.Text) 'gives the length of 512 increments of time
    
    lblTimeSlice.Caption = FORMAT(1000 * Ratio, "00.00") & " mSec"
    
    WaveF.SampleRate = Val(cboSampleRate.Text)
    
    pbxScale.Cls
    pbxScale.Line (0, 0)-(Ratio * pbxScale.ScaleWidth, pbxScale.ScaleHeight - 1), , BF
End Sub

Private Sub cboSampleRate_Click()
    cboSampleRate_Change
End Sub

Private Sub cboSampleRate_DropDown()
    cboSampleRate_Change
End Sub

Private Sub cmdOpen_Click()
    Dim FName As String
    
    With dlgCommon
        .FileName = ""
        .Filter = "Bitmap Images|*.bmp"
        .ShowOpen
        If .FileName = "" Then Exit Sub
        FName = .FileName
    End With
    
    pbxWave.Picture = LoadPicture(FName)
    
End Sub

Private Sub cmdPreview_Click()
    Dim x As Integer, y As Integer
    Dim Frequency As Double
    Dim Mult As Double
    
    cboSampleRate_Change
    
    'Length of audio data, does not include header length
    lblSize.Caption = (Val(cboBitsPer.Text) \ 8) * _
                      Val(txtTime.Text) * _
                      Val(cboSampleRate.Text) * _
                      WaveF.NumChannels & _
                      " bytes"
                      
    
    Frequency = Val(txtFreq.Text) 'get the frequency
    
    'erase picture
    pbxWave.Line (0, 0)-(pbxWave.ScaleWidth - 1, pbxWave.ScaleHeight - 1), pbxWave.BackColor, BF
    
    For x = 0 To pbxWave.ScaleWidth - 1
        
        Mult = x / Val(cboSampleRate.Text) 'current sample time
        
        Select Case WaveType
            Case Is = 0
                y = pbxWave.ScaleHeight - ((MAXVOL8 * scrVol.Value * 0.01) * _
                    Sin(2 * PI * Mult * Frequency) + 128)
                
            Case Is = 1
                y = pbxWave.ScaleHeight - ((MAXVOL8 * scrVol.Value * 0.01) * _
                    Cos(2 * PI * Mult * Frequency) + 128)
                    
            Case Is = 2
                y = pbxWave.ScaleHeight - ((MAXVOL8 * scrVol.Value * 0.01) * _
                    Cos(2 * PI * Mult * Frequency) + _
                    (MAXVOL8 * scrVol.Value * 0.002) * _
                    Sin(8 * PI * Mult * Frequency) + 128)
            
        End Select
        
        
        pbxWave.PSet (x, y)
    
    Next
    
End Sub

Private Sub cmdSave_Click()
    Dim FName As String
    
    With dlgCommon
        .FileName = ""
        .Filter = "Bitmap Images|*.bmp"
        .ShowSave
        If .FileName = "" Then Exit Sub
        FName = .FileName
    End With
    
    SavePicture pbxWave.Image, FName
    
End Sub

Private Sub cmdTrace_Click()
    Dim x As Integer, y As Integer
    Dim Col As Long
    
    For x = 0 To 511
    
        For y = 0 To pbxWave.ScaleHeight - 1
        
            Col = GetPixel(pbxWave.hdc, x, y)
            
            If Not Col = vbBlack Then
                
                Custom(x) = pbxWave.ScaleHeight - y - 128
                
                Exit For
                
            End If
        
        Next
        
    Next
    
    
    chkCustom.Enabled = True
End Sub

Private Sub cmdWrite_Click()
    On Error GoTo WRITEERR
    
    Dim FName As String
    Dim FNum As Integer
    Dim Frequency As Integer
    Dim MaxTime As Long
    Dim Mult As Double
    Dim Mag As Integer
    Dim ScaleMag As Integer
    Dim i As Long
    Dim n As Integer
    
    'File header setup
    WaveD.DataSize = (WaveF.BitsPerSample \ 8) * _
                      Val(txtTime.Text) * _
                      WaveF.SampleRate * _
                      WaveF.NumChannels
    
    WaveH.Size = WaveD.DataSize + 36
    
    WaveF.ByteRate = (WaveF.BitsPerSample \ 8) * WaveF.SampleRate * WaveF.NumChannels
    
    WaveF.BlockAlign = (WaveF.BitsPerSample \ 8) * WaveF.NumChannels
    
    Frequency = Val(txtFreq.Text)
    
    If WaveF.BitsPerSample = 8 Then
        ScaleMag = MAXVOL8
    ElseIf WaveF.BitsPerSample = 16 Then
        ScaleMag = MAXVOL16
    End If
    
    
    With dlgCommon
        .FileName = ""
        .Filter = "Wave Files|*.wav"
        .ShowSave
        If .FileName = "" Then Exit Sub
        FName = .FileName
    End With
    
    FNum = FreeFile
    
    frmMain.MousePointer = 11
    
    Open FName For Binary As FNum
    
    
    
        'Write wave file header info
        Put #FNum, , WaveH
        Put #FNum, , WaveF
        Put #FNum, , WaveD
        
        'Number of points in time (Samples) that we have, ignoring number of channels
        MaxTime = WaveF.SampleRate * Val(txtTime.Text)
        
        'Write actual data
        If chkCustom.Value = Checked Then
            
            '================Custom Waveform=====================================
            For i = 0 To MaxTime
                
                If WaveF.BitsPerSample = 8 Then
                    'This is already in 8 bit mode
                    Mag = Custom(n)
                    Mag = Mag + 128
                ElseIf WaveF.BitsPerSample = 16 Then
                    'convert 8 bit to 16 bit
                    Mag = (Custom(n) / 127) * ScaleMag
                End If
                
                If WaveF.NumChannels = 1 Then 'Write magnitude to single channel
                    
                    If WaveF.BitsPerSample = 8 Then
                        Put #FNum, , CByte(Mag) 'Write only 1 byte
                    Else
                        Put #FNum, , Mag 'Write 2 bytes (integer)
                    End If
                    
                ElseIf WaveF.NumChannels = 2 Then 'Write same magnitude to both channels
                    
                    If WaveF.BitsPerSample = 8 Then 'Write 1 byte for each channel
                        Put #FNum, , CByte(Mag)
                        Put #FNum, , CByte(Mag)
                    Else
                        Put #FNum, , Mag 'Write 2 bytes for each channel
                        Put #FNum, , Mag
                    End If
                    
                End If
                
                'Because we only have 512 samples viewable, we can only use that many
                'to create our custom waveform
                n = n + 1
                If n > 511 Then
                    n = 0
                End If
                
            Next
            '====================================================================
        Else
            '==================Sine Wave=========================================
            'Loop for as many samples as we have, ignoring number of channels
            MaxTime = WaveF.SampleRate * Val(txtTime.Text)
        
            For i = 0 To MaxTime
                
                Mult = i / MaxTime 'This is the point in time that is being written
                
                
                'Get the magnitude at the current time
                Select Case WaveType
                    Case Is = 0 'Sine waveform
                        Mag = (ScaleMag * scrVol.Value * 0.01) * _
                            Sin(2 * PI * Mult * Frequency)
                        
                    Case Is = 1 'Cosine waveform
                        Mag = (ScaleMag * scrVol.Value * 0.01) * _
                            Cos(2 * PI * Mult * Frequency)
                            
                    Case Is = 2 'Combination signal
                        Mag = (ScaleMag * scrVol.Value * 0.01) * _
                            Cos(2 * PI * Mult * Frequency) + _
                            (ScaleMag * scrVol.Value * 0.002) * _
                            Sin(8 * PI * Mult * Frequency)
                    
                End Select
                
                
                      
                If WaveF.BitsPerSample = 8 Then
                    Mag = Mag + 128
                End If
                
                If WaveF.NumChannels = 1 Then 'Write magnitude to single channel
                    
                    If WaveF.BitsPerSample = 8 Then
                        Put #FNum, , CByte(Mag)
                    Else
                        Put #FNum, , Mag
                    End If
                    
                ElseIf WaveF.NumChannels = 2 Then 'Write same magnitude to both channels
                    
                    If WaveF.BitsPerSample = 8 Then
                        Put #FNum, , CByte(Mag)
                        Put #FNum, , CByte(Mag)
                    Else
                        Put #FNum, , Mag
                        Put #FNum, , Mag
                    End If
                    
                End If
                
            Next
            '====================================================================
        
        End If
    
    
    
    Close #FNum
    
    frmMain.MousePointer = 0
    
    Exit Sub
    
WRITEERR:
    Close #FNum
    frmMain.MousePointer = 0
    MsgBox "Encountered an error when writing to file: " & FName, vbExclamation, "Write Error"
    
End Sub

Private Sub Form_Load()
    
    WaveH.ID = "RIFF"
    WaveH.FORMAT = "WAVE"
    
    With WaveF
        .ID = "fmt "
        .SUBSIZE = 16
        .AUDIOFORMAT = 1 'Wave PCM (non compressed)
        .NumChannels = 2
        .SampleRate = 44100
        .BitsPerSample = 16
    End With
    
    WaveD.ID = "data"
    
    WaveType = 0
    
End Sub

Private Sub optMono_Click()
    WaveF.NumChannels = 1
End Sub

Private Sub optSignal_Click(Index As Integer)
    WaveType = Index
End Sub

Private Sub optStereo_Click()
    WaveF.NumChannels = 2
End Sub

Private Sub pbxWave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawOn = 1
End Sub

Private Sub pbxWave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Mult As Double
    
    Mult = x / Val(cboSampleRate.Text)
    
    lblPos.Caption = FORMAT(1000 * Mult, "00.00") & " mSec"
    
    
    
    If DrawOn Then
        
        pbxWave.Line (x, 0)-(x, pbxWave.ScaleHeight - 1), pbxWave.BackColor
        
        pbxWave.PSet (x, y)
        
    End If
End Sub

Private Sub pbxWave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawOn = 0
End Sub

Private Sub scrVol_Change()
    lblVol.Caption = scrVol.Value & "%"
End Sub

Private Sub scrVol_Scroll()
    lblVol.Caption = scrVol.Value & "%"
End Sub
