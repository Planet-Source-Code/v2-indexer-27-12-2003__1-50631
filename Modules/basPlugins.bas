Attribute VB_Name = "basPlugins"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Type PlugDetails
    Name As String
    Version As String
    Description As String
End Type
Public Type Plug
    objMain As Object
    objInfo As Object
    objDetails As PlugDetails
    Name As String
End Type
Global sFileName As String
Global Plugs() As Plug


Function EnumPlugins()
On Error Resume Next
Dim TotalPlugs As Long
Dim nCtr As Long
Dim PlugName As String

Dim sName As String
Dim sVersion As String
Dim sDesc As String
Dim strFile As String


frmMain.lstFile.Path = App.Path & "\Plugins\"
frmMain.lstFile.Pattern = "*.PLUG"
frmMain.lstFile.Refresh
TotalPlugs = frmMain.lstFile.ListCount

If TotalPlugs <= 0 Then Exit Function
ReDim Preserve Plugs(TotalPlugs - 1) As Plug
'Register All Available Plugins To Avoid Any Error
For nCtr = 0 To TotalPlugs - 1
    DoEvents
    PlugName = Mid(frmMain.lstFile.List(nCtr), 1, Len(frmMain.lstFile.List(nCtr)) - 5)
    strFile = frmMain.lstFile.Path & "\" & frmMain.lstFile.List(nCtr)
    ShellExecute 0, "OPEN", "regsvr32", """" & strFile & """ /s", App.Path & "\plugins\", 0
Next


For nCtr = 0 To TotalPlugs - 1
    PlugName = Mid(frmMain.lstFile.List(nCtr), 1, Len(frmMain.lstFile.List(nCtr)) - 5)
    strFile = frmMain.lstFile.Path & "\" & frmMain.lstFile.List(nCtr)
    DoEvents
    Set Plugs(nCtr).objMain = CreateObject(PlugName & ".Main")
    Set Plugs(nCtr).objInfo = CreateObject(PlugName & ".Info")
    Plugs(nCtr).objInfo.Details sName, sVersion, sDesc
        
    If Not sName = PlugName Then
        FileCopy strFile, strFile & ".INVALID"
        Kill strFile
    Else
        Plugs(nCtr).objDetails.Name = sName
        Plugs(nCtr).objDetails.Version = sVersion
        Plugs(nCtr).objDetails.Description = sDesc
        Plugs(nCtr).Name = PlugName
    
        Load frmMain.mnuAboutPlug(frmMain.mnuAboutPlug.Count)
        frmMain.mnuAboutPlug(frmMain.mnuAboutPlug.Count - 1).Caption = PlugName
        frmMain.mnuAboutPlug(frmMain.mnuAboutPlug.Count - 1).Enabled = True
        frmMain.mnuAboutPlug(frmMain.mnuAboutPlug.Count - 1).Visible = True
        frmMain.cboHeader.AddItem PlugName
        sName = ""
        sVersion = ""
        sDesc = ""
    End If
    
Next
    frmMain.mnuAboutPlug(0).Visible = False
End Function


