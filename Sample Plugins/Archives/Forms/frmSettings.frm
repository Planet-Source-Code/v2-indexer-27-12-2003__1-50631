VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Archives Settings"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   2650
      TabIndex        =   7
      Top             =   1530
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Archives Format"
      Height          =   1365
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3500
      Begin VB.CheckBox chkFormat 
         Caption         =   "MSCOMPRESS 6.22"
         Height          =   285
         Index           =   5
         Left            =   1500
         TabIndex        =   6
         Top             =   960
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "MSCOMPRESS 5.0"
         Height          =   285
         Index           =   4
         Left            =   1500
         TabIndex        =   5
         Top             =   615
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "GZIP"
         Height          =   285
         Index           =   3
         Left            =   1500
         TabIndex        =   4
         Top             =   270
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "ARJ"
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   990
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "ZIP / RAR"
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   645
         Value           =   1  'Checked
         Width           =   1058
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "RAR"
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Value           =   1  'Checked
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
Dim nCtr As Byte

For nCtr = 0 To 5
    If chkFormat(nCtr).Value = 1 Then
        Select Case nCtr
          Case 0
                ReDim Preserve tmp(bCount) As FileType
                tmp(bCount).hdr = "Rar"                  'RAR
                tmp(bCount).StartOffset = 1
                tmp(bCount).Length = 3
                tmp(bCount).Type = CharString
            Case 1
                ReDim Preserve tmp(bCount) As FileType       'ZIP / JAR
                tmp(bCount).hdr = "PK"
                tmp(bCount).StartOffset = 1
                tmp(bCount).Length = 2
                tmp(bCount).Type = CharString
            Case 2
                ReDim Preserve tmp(bCount) As FileType       'ARJ
                tmp(bCount).hdr = "EA60"
                tmp(bCount).StartOffset = 1
                tmp(bCount).Length = 2
                tmp(bCount).Type = HexString
            Case 3
                ReDim Preserve tmp(bCount) As FileType       'GZIP
                tmp(bCount).hdr = "218B"
                tmp(bCount).StartOffset = 1
                tmp(bCount).Length = 2
                tmp(bCount).Type = HexString
            Case 4
                ReDim Preserve tmp(bCount) As FileType       'MSCOMPRESS 5.0
                tmp(bCount).hdr = "SZDD"
                tmp(bCount).StartOffset = 1
                tmp(bCount).Length = 3
                tmp(bCount).Type = CharString
            Case 5
                ReDim Preserve tmp(bCount) As FileType       'MSCOMPRESS 6.22
                tmp(bCount).hdr = "KWAJ"
                tmp(bCount).StartOffset = 1
                tmp(bCount).Length = 3
                tmp(bCount).Type = CharString
        End Select
        bCount = bCount + 1
    End If
Next
Unload Me
End Sub
