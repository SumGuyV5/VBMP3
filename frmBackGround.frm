VERSION 5.00
Begin VB.Form frmBackGround 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox filBackGround 
      Height          =   2040
      Left            =   120
      Pattern         =   "*.gif"
      TabIndex        =   5
      Top             =   2400
      Width           =   4335
   End
   Begin VB.DriveListBox drvBackGround 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.DirListBox dirBackGround 
      Height          =   1890
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1320
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   4560
      Width           =   1320
   End
End
Attribute VB_Name = "frmBackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
    strBackGround = dirBackGround.Path & "\" & filBackGround.FileName
    Open App.Path & "/Settings.dat" For Output As #1
        Write #1, strMP3Path
        Write #1, strBackGround
    Close #1
    
    Exit Sub
ErrorHandler:
    MsgBox "Unable to Save Path!!", vbCritical, "Error"
End Sub

Private Sub dirBackGround_Change()
    filBackGround.Path = dirBackGround.Path
End Sub

Private Sub drvBackGround_Change()
    dirBackGround.Path = drvBackGround.Drive
End Sub
