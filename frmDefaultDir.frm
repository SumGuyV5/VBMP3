VERSION 5.00
Begin VB.Form frmDefaultDir 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1320
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1320
   End
   Begin VB.DirListBox dirDefault 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.DriveListBox drvDefault 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmDefaultDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2003 Richard W. Allen
'Program Name:  VB MP3
'Author:        Richard W. Allen
'Version:       V1.0
'Date Started:  March 27, 2003
'Date End:      April 8, 2003
'"VB MP3!" Copyright (C) 2003 Richard W. Allen  "VB MP3!" Comes with ABSOLUTELY NO WARRANTY;
'VB MP3! is licensed under the GNU GENERAL PUBLIC LICENSE Version 2.
'for details see the license.txt include with this program.
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Private Sub cmdReset_Click()
    drvDefault.Drive = App.Path
    dirDefault.Path = App.Path
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
    strMP3Path = dirDefault.Path
    Open App.Path & "/Settings.dat" For Output As #1
        Write #1, strMP3Path
    Close #1
    frmMain.drvDriveSelect.Drive = strMP3Path
    frmMain.dirDirectoryListBox.Path = strMP3Path
    Exit Sub
ErrorHandler:
    MsgBox "Unable to Save Path!!", vbCritical, "Error"
End Sub

Private Sub drvDefault_Change()
    dirDefault.Path = drvDefault.Drive
End Sub

Private Sub Form_Load()
    drvDefault.Drive = App.Path
    dirDefault.Path = App.Path
End Sub
