VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Function GetType(nCount As Integer) As FileType()

Dim tmp(5) As FileType
            'RAR
                tmp(0).hdr = "Rar"
                tmp(0).StartOffset = 1
                tmp(0).Length = 3
                tmp(0).Type = CharString
            
            'ZIP / JAR
                tmp(1).hdr = "PK"
                tmp(1).StartOffset = 1
                tmp(1).Length = 2
                tmp(1).Type = CharString
            
            'ARJ
                tmp(2).hdr = "EA60"
                tmp(2).StartOffset = 1
                tmp(2).Length = 2
                tmp(2).Type = HexString
            
            'GZIP
                tmp(3).hdr = "218B"
                tmp(3).StartOffset = 1
                tmp(3).Length = 2
                tmp(3).Type = HexString
            
            'MSCOMPRESS 5.0
                tmp(4).hdr = "SZDD"
                tmp(4).StartOffset = 1
                tmp(4).Length = 3
                tmp(4).Type = CharString
            
            'MSCOMPRESS 6.22
                tmp(5).hdr = "KWAJ"
                tmp(5).StartOffset = 1
                tmp(5).Length = 3
                tmp(5).Type = CharString
nCount = 6
GetType = tmp
End Function

Function Init()
    DoEvents
End Function
