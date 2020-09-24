VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indexer"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11730
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   782
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgLstSmall 
      Left            =   9090
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2294
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":282E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search In"
      Height          =   1905
      Left            =   150
      TabIndex        =   12
      Top             =   60
      Width           =   2475
      Begin VB.DirListBox lstDir 
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2235
      End
      Begin VB.DriveListBox lstDrive 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   2265
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search For"
      Height          =   1905
      Left            =   2640
      TabIndex        =   2
      Top             =   60
      Width           =   4785
      Begin VB.ComboBox cboHeader 
         Height          =   315
         ItemData        =   "frmMain.frx":3362
         Left            =   1110
         List            =   "frmMain.frx":337B
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Width           =   2300
      End
      Begin VB.ComboBox txtPattern 
         Height          =   315
         ItemData        =   "frmMain.frx":33F1
         Left            =   1110
         List            =   "frmMain.frx":3407
         TabIndex        =   8
         Text            =   "All Files"
         Top             =   270
         Width           =   2300
      End
      Begin VB.ComboBox cboOptions 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMain.frx":345C
         Left            =   1110
         List            =   "frmMain.frx":3466
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   660
         Width           =   2300
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         ItemData        =   "frmMain.frx":3489
         Left            =   1110
         List            =   "frmMain.frx":34A5
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1050
         Width           =   2300
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save To &File"
         Enabled         =   0   'False
         Height          =   400
         Left            =   3510
         TabIndex        =   5
         Top             =   1350
         Width           =   1095
      End
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "&Quit"
         Height          =   400
         Left            =   3510
         TabIndex        =   4
         Top             =   825
         Width           =   1095
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   400
         Left            =   3510
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Header:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Text:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   330
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Size:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   1110
         Width           =   630
      End
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   5760
      Top             =   1470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox lstFile 
      Height          =   480
      Hidden          =   -1  'True
      Left            =   9870
      System          =   -1  'True
      TabIndex        =   0
      Top             =   4980
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSComctlLib.ListView lstResult 
      Height          =   6300
      Left            =   150
      TabIndex        =   17
      Top             =   2070
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   11113
      SortKey         =   1
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "imgLstBig"
      SmallIcons      =   "imgLstSmall"
      ColHdrIcons     =   "imgLstSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Extension"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLstBig 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":351E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indexer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   8400
      Width           =   705
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Search"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Search"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLastSearch 
         Caption         =   "Remember &Last Search"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewIcons 
         Caption         =   "As Icons"
      End
      Begin VB.Menu mnuViewSmallIcons 
         Caption         =   "As Small Icons"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "As List"
      End
      Begin VB.Menu mnuViewDetails 
         Caption         =   "As Details"
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutPlugins 
         Caption         =   "About Plugins"
         Begin VB.Menu mnuAboutPlug 
            Caption         =   "No Plugin Found"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Open &File"
      End
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "Open &Containing Folder"
      End
      Begin VB.Menu mnuCopyPath 
         Caption         =   "Copy Complete Path Into Clipboard"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastHeader As Byte
Private Sub btnCancel_Click()

On Error GoTo errHandler
    
    Unload Me
    End

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Public Sub btnSave_Click()

On Error GoTo errHandler
    
Dim fName As String
Dim fNum As Integer
Dim nCtr As Long

    If lstResult.ListItems.Count < 1 Then MsgBox "Nothing To Save", vbInformation: Exit Sub
    
    With cmd
        .Flags = 4
        .Filter = "Text Files(*.txt)|*.txt|Comma Seperated(*.csv)|*.csv"
        .ShowSave
        fName = .FileName
    End With
    
    Select Case cmd.FilterIndex
    Case 1
        If Not Right(fName, 4) = ".txt" Then fName = fName & ".txt"
        If Len(fName) <= 0 Then Exit Sub
        fNum = FreeFile
        Open fName For Output As fNum
                For nCtr = 1 To lstResult.ListItems.Count
                    Print #fNum, lstResult.ListItems(nCtr).ListSubItems(1).Text & "\" & lstResult.ListItems(nCtr).Text
                Next nCtr
        Close fNum
    Case 2
        If Not Right(fName, 4) = ".csv" Then fName = fName & ".csv"
        If Len(fName) <= 0 Then Exit Sub
        fNum = FreeFile
        Open fName For Output As fNum
                For nCtr = 1 To lstResult.ListItems.Count
                    Print #fNum, """" & lstResult.ListItems(nCtr).Text & """,""" & lstResult.ListItems(nCtr).SubItems(1) & """,""" & lstResult.ListItems(nCtr).SubItems(2) & """,""" & lstResult.ListItems(nCtr).SubItems(3) & """"
                Next nCtr
        Close fNum
    End Select
    
    
errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub btnSearch_Click()

On Error GoTo errHandler
    
    Dim sec As Long
    Dim strSize As String
    
If btnSearch.Caption = "&Search" Then
    If lstResult.ListItems.Count > 0 Then
        If MsgBox("This Will Clear The Result Of Last Search , Are You Sure You Want To Continue.", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    btnSearch.Caption = "&Stop"
    boolStop = False
    FirstFolder = lstDir.Path
    lstResult.ListItems.Clear
    lstResult.Sorted = False
    lblFile = "Searching..."
    TotalFolder = 0
    TotalFile = 0
    TotalSize = 0
    t1 = GetTickCount
    btnSave.Enabled = False
    
    If cboOptions.ListIndex = 1 Then
        Search_File lstDir.Path, "*" & SearchPattern & "*", cboSize.ListIndex
    Else
        Search_File lstDir.Path, SearchPattern, cboSize.ListIndex
    End If
    
    t2 = GetTickCount
    sec = t2 - t1
    strSize = TotalSize & " Bytes."
    If TotalSize >= 1024 And TotalSize < 1048576 Then strSize = Round(TotalSize / 1024, 2) & " KiloBytes."
    If TotalSize >= 1048576 And TotalSize < 1073741824 Then strSize = Round(TotalSize / 1048576, 2) & " MegaBytes."
    If TotalSize >= 1073741824 And TotalSize < 1099511627776# Then strSize = Round(TotalSize / 1073741824, 2) & " GegaBytes."
    
    lblFile.Caption = "File(s): " & TotalFile & Space(5) & strSize & Space(5) & "Time Taken: " & Int(sec / 1000) & " Seconds."
    btnSearch.Caption = "&Search"
Else
    btnSearch.Caption = "&Search"
    boolStop = True
    lblFile.Caption = "File(s): " & TotalFile & Space(5) & strSize & Space(5) & "Time Taken: " & Int(sec / 1000) & " Seconds."
End If
    
    Me.Caption = "Indexer"
    If lstResult.ListItems.Count > 0 Then
        btnSave.Enabled = True
        mnuSave.Enabled = True
        mnuClear.Enabled = True
    Else
        btnSave.Enabled = False
        mnuSave.Enabled = False
        mnuClear.Enabled = False
    End If

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub




Private Sub cboHeader_Click()
    If cboHeader.ListIndex = 6 Then cboHeader.ListIndex = 0
End Sub

Private Sub Form_Load()

On Error GoTo errHandler
    Hide
    EnumPlugins
    LastHeader = 1
    lstResult.ColumnHeaders(1).Icon = 1
    cboOptions.ListIndex = 0
    cboSize.ListIndex = 0
    lblFile.Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    mnuLastSearch.Checked = GetSetting(App.ProductName, App.Major & "." & App.Minor & "\Settings", "Save Last Search", True)
If mnuLastSearch.Checked = True Then
    lstDrive.Drive = GetSetting(App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "Folder", CurDir)
    lstDir.Path = GetSetting(App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "Folder", CurDir)
    txtPattern.Text = GetSetting(App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "Search Text", "All Files")
    txtPattern.ListIndex = GetSetting(App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "File Types", 0)
    cboOptions.ListIndex = GetSetting(App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "Options", 0)
    cboSize.ListIndex = GetSetting(App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "File Size", 0)
    cboHeader.ListIndex = GetSetting(App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "File Header", 0)
End If
    Show
    
SetFileTypes

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub Form_Resize()
    Frame1.Move 10, 3, (ScaleWidth - Frame2.Width - 30)
    Frame2.Move Frame1.Left + Frame1.Width + 10, 3
    lstDrive.Move 100, 250, (Frame1.Width * 15) - 200
    lstDir.Move lstDrive.Left, lstDrive.Top + lstDrive.Height + 5, lstDrive.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo errHandler
    
     SaveSetting App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "Folder", lstDir.Path
     SaveSetting App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "Search Text", txtPattern.Text
     SaveSetting App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "File Types", txtPattern.ListIndex
     SaveSetting App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "Options", cboOptions.ListIndex
     SaveSetting App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "File Size", cboSize.ListIndex
     SaveSetting App.ProductName, App.Major & "." & App.Minor & "\Settings", "Save Last Search", mnuLastSearch.Checked
     SaveSetting App.ProductName, App.Major & "." & App.Minor & "\Settings\LastSearch", "File Header", cboHeader.ListIndex

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub Frame1_DblClick()
On Error GoTo errHandler
   
   lstDir.Path = InputBox("Enter The Path To Change Directory To:")
   lstFile.Path = lstDir.Path
   lstDrive.Drive = lstDir.Path

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub lstDrive_Change()

On Error GoTo errHandler
    
    lstDir.Path = lstDrive.Drive

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub lstDir_Change()

On Error GoTo errHandler
    
    lstFile.Path = lstDir.Path

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub lstResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstResult.SortKey = ColumnHeader.Index - 1
    lstResult.ColumnHeaders(LastHeader).Icon = 3
    
    If ColumnHeader.Index = LastHeader Then
        If lstResult.SortOrder = lvwAscending Then
            lstResult.SortOrder = lvwDescending
            ColumnHeader.Icon = 2
        Else
            lstResult.SortOrder = lvwAscending
            ColumnHeader.Icon = 1
        End If
    Else
            lstResult.SortOrder = lvwAscending
            ColumnHeader.Icon = 1
    End If
    
    lstResult.Sorted = True
    LastHeader = ColumnHeader.Index
    
End Sub

Private Sub lstResult_DblClick()

On Error GoTo errHandler
   
   If lstResult.ListItems.Count < 1 Then Exit Sub
    ShellExecute 0, "OPEN", lstResult.SelectedItem.ListSubItems(1).Text & "\" & lstResult.SelectedItem.Text, "", lstResult.SelectedItem.ListSubItems(1).Text, 1

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub


Private Sub lstResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstResult.ListItems.Count > 0 Then
        If Button = 2 Then
            PopupMenu mnuPopUp
        End If
    End If
End Sub

Private Sub mnuAbout_Click()

On Error GoTo errHandler
    
    MsgBox "Indexer - [ Version: " & App.Major & "." & App.Minor & "." & App.Revision & " ]" & vbCr & "Developer: Vikas Verma" & vbCr & "Email: v2softwares@yahoo.com" & vbCr & "Website: http://www.geocities.com/v2softwares" & vbCr & vbCr & _
           "This Small Utility Is Handy To Create Text Files Containing List Of Files On The HardDisk Or Any Other Removable Drive For Offline Search Options.", vbInformation, "About Indexer v" & App.Major & "." & App.Minor & "." & App.Revision

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub mnuAboutPlug_Click(Index As Integer)
    Plugs(Index - 1).objInfo.About
End Sub

Private Sub mnuClear_Click()

On Error GoTo errHandler
    
    If lstResult.ListItems.Count > 0 Then
        If MsgBox("This Will Clear The Result Of Last Search , Are You Sure You Want To Continue.", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        Else
            lstResult.ListItems.Clear
            btnSave.Enabled = False
            
        End If
    End If

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub mnuCopyPath_Click()
    Clipboard.SetText lstResult.SelectedItem.ListSubItems(1).Text & "\" & lstResult.SelectedItem.Text
End Sub

Private Sub mnuExit_Click()

On Error GoTo errHandler
    
    btnCancel_Click

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub mnuLastSearch_Click()

On Error GoTo errHandler
    
    mnuLastSearch.Checked = Not mnuLastSearch.Checked

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub


Private Sub mnuOpenFile_Click()
    ShellExecute 0, "OPEN", lstResult.SelectedItem.ListSubItems(1).Text & "\" & lstResult.SelectedItem.Text, "", lstResult.SelectedItem.ListSubItems(1).Text, 1
End Sub

Private Sub mnuOpenFolder_Click()
    ShellExecute 0, "OPEN", lstResult.SelectedItem.ListSubItems(1).Text, "", lstResult.SelectedItem.ListSubItems(1).Text, 1
End Sub

Private Sub mnuSave_Click()

On Error GoTo errHandler
    
    btnSave_Click

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub mnuViewDetails_Click()
    lstResult.View = lvwReport
End Sub

Private Sub mnuViewIcons_Click()
    lstResult.View = lvwIcon
End Sub

Private Sub mnuViewList_Click()
    lstResult.View = lvwList
End Sub

Private Sub mnuViewRefresh_Click()
    lstResult.Refresh
End Sub

Private Sub mnuViewSmallIcons_Click()
    lstResult.View = lvwSmallIcon
End Sub



Private Sub txtPattern_Change()

On Error GoTo errHandler
    
    txtPattern_Click

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub

Private Sub txtPattern_Click()

On Error GoTo errHandler
    
Select Case txtPattern.ListIndex
Case 0
    SearchPattern = "*.*"
    cboOptions.Enabled = False
Case 1
    SearchPattern = "*.jpg;*.bmp;*.wmf;*.gif;*.psd;*.tif;*.png;*.dib;*.ico;*.cur"
    cboOptions.Enabled = False
Case 2
    SearchPattern = "*.wav;*.mp1;*.mp2;*.mp3;*.wma;*.aiff;*.ra"
    cboOptions.Enabled = False
Case 3
    SearchPattern = "*.dat;*.avi;*.mpg;*.mpeg;*.wmv;*.rm;*.mov;*.asx;*.asf;*mpa"
    cboOptions.Enabled = False
Case 4
    SearchPattern = "*.txt;*.wri;*.rtf;*.doc;*.pdf"
    cboOptions.Enabled = False
Case 5
    SearchPattern = "*.exe;*.com;*.sys;*.dll;*.ocx;*.cpl;*.drv;*.ax;*.vxd"
    cboOptions.Enabled = False
Case Else
    SearchPattern = txtPattern.Text
    cboOptions.Enabled = True
    cboOptions.ListIndex = 1
End Select

errHandler:
If Not Err.Number = 0 Then
    MsgBox "Error: " & Err.Description & " ( " & Err.Number & " )", vbExclamation, "Error Source: " & Err.Source
End If

End Sub
