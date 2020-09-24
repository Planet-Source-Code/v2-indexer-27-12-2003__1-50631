Attribute VB_Name = "basDeclarations"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Global SearchPattern As String
Global t1 As Long, t2 As Long
Global TotalFolder As Long
Global TotalFile As Long
Global TotalFileFound As Long
Global boolStop As Boolean
Global FirstFolder As String

Global Const MB_5 As Long = 5242880
Global Const MB_10 As Long = 10485760
Global Const MB_50 As Long = 52428800
Global Const MB_100 As Long = 104857600
Global Const GB_1 As Long = 1073741824

Global TotalSize As Double


Public Function Search_File(SFolder As String, Optional Pattern As String = "*.*", Optional FileSize As Integer = -1)

On Error GoTo errHandler

Dim FileHeader As String * 20
Dim SearchComplete  As Boolean
Dim LastFolder      As String
Dim FolderCount     As Long
Dim FileCount     As Long
Dim nCtr As Long
Dim nCtr1 As Long
Dim n As Integer
Dim strFile As String
Dim lngLen As Long
Dim tmpFile() As String
Dim FileNum As Integer
Dim sSize As String
Dim sFile() As FileType
Dim Ext As String
Dim tt1 As Long
Dim tt2 As Long
Dim tm As Long

    LastFolder = SFolder
    frmMain.lstFile.Pattern = Pattern
    frmMain.lstDir.Path = LastFolder
    
    FolderCount = frmMain.lstDir.ListCount
    FileCount = frmMain.lstFile.ListCount
    TotalFileFound = TotalFileFound + FileCount
    tt1 = GetTickCount
    
    For nCtr = 0 To FolderCount - 1
        DoEvents
        If boolStop = True Then
            Exit Function
        End If
        Search_File frmMain.lstDir.List(nCtr), Pattern, FileSize
        frmMain.lstDir.Path = LastFolder
    Next nCtr
        
    frmMain.Caption = "Indexer - [ " & LastFolder & " ]"
        
    SetFilterType sFile, n
    For nCtr1 = 0 To FileCount - 1
        
        'frmMain.lblFile.Caption = "Total File(s): " & TotalFile & " , Time: " & Int((GetTickCount - t1) / 1000) & " Seconds."
        tt2 = GetTickCount
        tm = Int(tt2 - tt1) / 1000
        If tm <= 0 Then tm = 1
        frmMain.lblFile.Caption = "Seaching..... Speed: " & Int(TotalFileFound / tm) & " Files Per Second."
        
        If boolStop = True Then
            Exit Function
        End If
        strFile = frmMain.lstFile.Path & "\" & frmMain.lstFile.List(nCtr1)
        lngLen = FileLen(strFile)
        FileNum = FreeFile
        
        Open strFile For Binary As FileNum
            Get FileNum, , FileHeader
        Close FileNum
        If FilterType(FileHeader, sFile, n) = True Then
            Select Case FileSize
                Case 0
                    If FilterSize(strFile, 0, 0, 0, False, True, False) = True Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
                Case 1
                    If FilterSize(strFile, 0, 0, 0, True, False, False) Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
                Case 2
                    If FilterSize(strFile, 0, MB_5, 0, False, False, False) Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
                Case 3
                    If FilterSize(strFile, MB_5, MB_10, 0, False, False, False) Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
                Case 4
                    If FilterSize(strFile, MB_10, MB_50, 0, False, False, False) Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
                Case 5
                    If FilterSize(strFile, MB_50, MB_100, 0, False, False, False) Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
                Case 6
                    If FilterSize(strFile, MB_100, GB_1, 0, False, False, False) Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
                Case 7
                    If FilterSize(strFile, GB_1, 0, 0, False, False, True) Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
                Case Else
                    If FilterSize(strFile, 0, 0, 0, False, True, False) Then
                        frmMain.lstResult.ListItems.Add , , Mid(strFile, InStrRev(strFile, "\") + 1, Len(strFile) - InStrRev(strFile, "\")), 1, 4
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 1, , Mid(strFile, 1, InStrRev(strFile, "\") - 1)
                        If InStrRev(strFile, ".") > 1 Then
                            Ext = UCase(Mid(strFile, InStrRev(strFile, ".") + 1, Len(strFile) - InStrRev(strFile, ".")))
                        Else
                            Ext = "No Extension"
                        End If
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 2, , Ext
                        ConvertSize lngLen, sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ListSubItems.Add 3, , sSize
                        frmMain.lstResult.ListItems(frmMain.lstResult.ListItems.Count).ToolTipText = strFile
                        TotalFile = TotalFile + 1
                        TotalSize = TotalSize + lngLen
                    End If
            End Select
        End If
    Next nCtr1
    

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Function


Function ConvertSize(lngByte As Long, SizeString As String, Optional nCount As Integer = 0, Optional Done As Boolean = False) As Long
    
    If lngByte > 1024 Then
        ConvertSize lngByte / 1024, SizeString, nCount + 1, Done
    End If
    If Done = True Then Exit Function
    
    Select Case nCount
        Case 0
                SizeString = lngByte & " Byte"
                Done = True
        Case 1
                SizeString = lngByte & " KB"
                Done = True
        Case 2
                SizeString = lngByte & " MB"
                Done = True
        Case 3
                SizeString = lngByte & " GB"
                Done = True
        Case 4
                SizeString = lngByte & " TB"
                Done = True
        Case Else
                SizeString = "Very Huge File"
                Done = True
    End Select
    
End Function

Function SetFilterType(sFilterType() As FileType, nCount As Integer)
    Select Case frmMain.cboHeader.ListIndex
        Case 1
            ReDim CurrentType(8) As FileType
            CurrentType(0) = JPEG
            CurrentType(1) = BMP
            CurrentType(2) = WMF
            CurrentType(3) = GIF
            CurrentType(4) = PSD
            CurrentType(5) = TIF
            CurrentType(6) = TIFF
            CurrentType(7) = PNG
            CurrentType(8) = ANI
            nCount = 9
        Case 2
            ReDim CurrentType(4) As FileType
            CurrentType(0) = AVI
            CurrentType(1) = MOV
            CurrentType(2) = DAT
            CurrentType(3) = MPG
            CurrentType(4) = WMA
            nCount = 5
        Case 3
            ReDim CurrentType(0) As FileType
            CurrentType(0) = WAV
            nCount = 1
        Case 4
            ReDim CurrentType(0) As FileType
            CurrentType(0) = EXE
            nCount = 1
        Case 5
            ReDim CurrentType(0) As FileType
            CurrentType(0) = ZIP
            nCount = 1
        Case 6
            ReDim CurrentType(0) As FileType
            CurrentType(0) = ZIP
            nCount = 1
        Case Else
            If frmMain.cboHeader.ListIndex > 6 Then
                Call Plugs(frmMain.cboHeader.ListIndex - 7).objMain.Init
                CurrentType = Plugs(frmMain.cboHeader.ListIndex - 7).objMain.GetType(nCount)
            End If
    End Select
    sFilterType = CurrentType
End Function

