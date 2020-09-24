Attribute VB_Name = "basFilter"
Function FilterType(strSource As String, strParameter() As FileType, nCount As Integer) As Boolean
Dim strMatch As String
Dim nStart As Integer
Dim nLen As Integer
Dim nCtr As Integer

If nCount < 1 Then FilterType = True: Exit Function

For nCtr = 0 To nCount - 1
    If strParameter(nCtr).Type = HexString Then
        strMatch = HexToStr(strParameter(nCtr).hdr)
    Else
        strMatch = strParameter(nCtr).hdr
    End If
    nStart = strParameter(nCtr).StartOffset
    nLen = CInt(strParameter(nCtr).Length)
    
        If Mid(strSource, nStart, nLen) = strMatch Then
            FilterType = True
            Exit Function
        Else
            FilterType = False
        End If
Next
End Function

Function FilterSize(strSource As String, lngMinSize As Long, lngMaxSize As Long, Optional ExactSize As Long = 0, Optional DoExact As Boolean = False, Optional AlwaysTrue As Boolean = False, Optional NoMaxLimit As Boolean = False) As Boolean
    If AlwaysTrue = True Then FilterSize = True: Exit Function
    
    If NoMaxLimit Then
        If FileLen(strSource) >= lngMinSize Then
            FilterSize = True
        Else
            FilterSize = False
        End If
        Exit Function
    End If
    
    If DoExact = True Then
        If FileLen(strSource) = ExactSize Then
            FilterSize = True
        Else
            FilterSize = False
        End If
    Else
        If FileLen(strSource) >= lngMinSize And FileLen(strSource) <= lngMaxSize Then
            FilterSize = True
        Else
            FilterSize = False
        End If
    End If
End Function


