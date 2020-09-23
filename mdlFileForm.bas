Attribute VB_Name = "mdlFileForm"
Option Explicit
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Global Const PI As Double = 3.14159     'mmm... pi
Global Const DEGTORAD As Double = PI / 180
Global Const MAXVOL16 As Integer = 30000
Global Const MAXVOL8 As Integer = 127

Type WAVEHEADER                         'Chunk
    ID As String * 4                    'the beginning of the file
    Size As Long                        'Size of file following this info
                                        '=36 + SubChunk2Size, or:
                                        '4 + (8 + SubChunk1Size) + (8 + SubChunk2Size)
    FORMAT As String * 4                'Tells that this file is the WAVE subformat of the RIFF spec
End Type

Type FORMATHEADER                       'Chunk1
    ID As String * 4
    SUBSIZE As Long                     'the remaining size of this subchunk 16 for WAVE PCM format
    AUDIOFORMAT As Integer              'PCM = 1 Linear quantization
    NumChannels As Integer
    SampleRate As Long
    ByteRate As Long                    'SampleRate * NumChannels * BitsPerSample/8
    BlockAlign As Integer               'NumChannels * BitsPerSample/8
    BitsPerSample As Integer
End Type

Type DATAHEADER                         'Chunk2
    ID As String * 4
    DataSize As Long                    '=NumSamples * NumChannels * BitsPerSample/8
    'Data() as integer                  'Pointer to sound data
End Type

Global Ratio As Double
