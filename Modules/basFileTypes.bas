Attribute VB_Name = "basFileTypes"
Type FileType
    hdr As String
    StartOffset As Integer
    Length As Integer
    Type As enumType
End Type

Enum enumType
    CharString = 0
    HexString = 1
End Enum

'Picture Files
Global JPEG As FileType
Global BMP As FileType
Global WMF As FileType
Global GIF As FileType
Global PSD As FileType
Global TIF As FileType
Global TIFF As FileType
Global PNG As FileType
Global ANI As FileType

'Video Files
Global AVI As FileType
Global MOV As FileType
Global DAT As FileType
Global MPG As FileType
Global WMA As FileType
Global RAM As FileType      'TO DO
Global RM As FileType       'TO DO

'Audio Files
Global AIFF As FileType     'TO DO
Global WAV As FileType
Global MP3 As FileType      'TO DO
Global SND As FileType      'TO DO
Global RA As FileType       'TO DO

'Text Documents
'TO DO

'Executables
Global EXE As FileType

'Archives
Global ZIP As FileType
Global CAB As FileType      'TO DO
Global RAR As FileType      'TO DO
Global ARC As FileType      'TO DO
Global TAR As FileType      'TO DO
Global ACE As FileType      'TO DO

Global CurrentType() As FileType


Function SetFileTypes()
    SetPictureFiles
    SetVideoFiles
    SetAudioFiles
    SetDocFiles
    SetExeFiles
    SetArchivesFiles
End Function

Function SetPictureFiles()
With JPEG
    .hdr = "JFIF"
    .StartOffset = 7
    .Length = 4
    .Type = CharString
End With

With BMP
    .hdr = "BM"
    .StartOffset = 1
    .Length = 2
    .Type = CharString
End With

With WMF
    .hdr = "D7CDC69A"
    .StartOffset = 1
    .Length = 4
    .Type = HexString
End With

With GIF
    .hdr = "GIF"
    .StartOffset = 1
    .Length = 3
    .Type = CharString
End With

With PSD
    .hdr = "8BPS"
    .StartOffset = 1
    .Length = 4
    .Type = CharString
End With

With TIF
    .hdr = "II"
    .StartOffset = 1
    .Length = 2
    .Type = CharString
End With

With TIFF
    .hdr = "MM"
    .StartOffset = 1
    .Length = 2
    .Type = CharString
End With

With PNG
    .hdr = "png"
    .StartOffset = 2
    .Length = 3
    .Type = CharString
End With

With ANI
    .hdr = "ACON"
    .StartOffset = 9
    .Length = 4
    .Type = CharString
End With

End Function

Function SetVideoFiles()
With AVI
    .hdr = "AVI"
    .StartOffset = 9
    .Length = 3
    .Type = CharString
End With

With MOV
    .hdr = "moov"
    .StartOffset = 5
    .Length = 4
    .Type = CharString
End With

With DAT
    .hdr = "CDXA"
    .StartOffset = 9
    .Length = 4
    .Type = CharString
End With

With MPG
    .hdr = "000001BA21000100"
    .StartOffset = 1
    .Length = 8
    .Type = HexString
End With

With WMA
    .hdr = "3026B2758E66CF11"
    .StartOffset = 1
    .Length = 8
    .Type = HexString
End With
End Function

Function SetAudioFiles()
With WAV
    .hdr = "WAVE"
    .StartOffset = 9
    .Length = 4
    .Type = CharString
End With
End Function

Function SetDocFiles()
End Function

Function SetExeFiles()
With EXE
    .hdr = "MZ"
    .StartOffset = 1
    .Length = 2
    .Type = CharString
End With
End Function

Function SetArchivesFiles()
With ZIP
    .hdr = "PK"
    .StartOffset = 1
    .Length = 2
    .Type = CharString
End With
End Function

