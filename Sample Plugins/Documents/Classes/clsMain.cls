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
Dim tmp(0) As FileType

    With tmp(0)
        .hdr = "%PDF"
        .StartOffset = 1
        .Length = 4
        .Type = CharString
    End With
    
    nCount = 1
    GetType = tmp
End Function

Function Init()
    DoEvents
End Function

