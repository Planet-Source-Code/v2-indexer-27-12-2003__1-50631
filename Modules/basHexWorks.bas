Attribute VB_Name = "basHexWorks"
Function HexToDec(HexStr As String) As Long
Dim strlen As Integer
Dim Ctr As Integer
Dim Word As String * 1
Dim lngTemp As Long

    strlen = Len(HexStr)
    
    HexStr = UCase(HexStr)
    
    For Ctr = strlen To 1 Step -1
        
        Word = Left$(HexStr, 1)
        
        HexStr = Mid(HexStr, 2, Len(HexStr) - 1)
        
        Select Case Word
            
            Case "0"
                lngTemp = lngTemp
            Case "1"
                lngTemp = (16 ^ (Ctr - 1)) * 1 + lngTemp
            Case "2"
                lngTemp = (16 ^ (Ctr - 1)) * 2 + lngTemp
            Case "3"
                lngTemp = (16 ^ (Ctr - 1)) * 3 + lngTemp
            Case "4"
                lngTemp = (16 ^ (Ctr - 1)) * 4 + lngTemp
            Case "5"
                lngTemp = (16 ^ (Ctr - 1)) * 5 + lngTemp
            Case "6"
                lngTemp = (16 ^ (Ctr - 1)) * 6 + lngTemp
            Case "7"
                lngTemp = (16 ^ (Ctr - 1)) * 7 + lngTemp
            Case "8"
                lngTemp = (16 ^ (Ctr - 1)) * 8 + lngTemp
            Case "9"
                lngTemp = (16 ^ (Ctr - 1)) * 9 + lngTemp
            Case "A"
                lngTemp = (16 ^ (Ctr - 1)) * 10 + lngTemp
            Case "B"
                lngTemp = (16 ^ (Ctr - 1)) * 11 + lngTemp
            Case "C"
                lngTemp = (16 ^ (Ctr - 1)) * 12 + lngTemp
            Case "D"
                lngTemp = (16 ^ (Ctr - 1)) * 13 + lngTemp
            Case "E"
                lngTemp = (16 ^ (Ctr - 1)) * 14 + lngTemp
            Case "F"
                lngTemp = (16 ^ (Ctr - 1)) * 15 + lngTemp
            
            Case Else
        End Select
        
    Next
    HexToDec = lngTemp
    
End Function

Function HexToStr(HexStr As String) As String
Dim nCtr As Long
Dim nCount As Long
Dim tmpBuffer As String
nCount = Len(HexStr)
If Not (nCount Mod 2) = 0 Then Exit Function

For nCtr = 1 To nCount Step 2
    tmpBuffer = tmpBuffer & Chr(HexToDec(Mid(HexStr, nCtr, 2)))
Next
    HexToStr = tmpBuffer
End Function
