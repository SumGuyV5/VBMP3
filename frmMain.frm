VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   6375
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1
      Left            =   7800
      Top             =   1320
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.HScrollBar hsbSpeed 
      Height          =   255
      LargeChange     =   50
      Left            =   4440
      Max             =   200
      Min             =   50
      TabIndex        =   9
      Top             =   2520
      Value           =   50
      Width           =   3615
   End
   Begin VB.HScrollBar hsbBalance 
      Height          =   255
      LargeChange     =   500
      Left            =   4440
      Max             =   5000
      Min             =   -5000
      SmallChange     =   100
      TabIndex        =   10
      Top             =   3600
      Width           =   3615
   End
   Begin VB.HScrollBar hsbVolume 
      Height          =   255
      Left            =   4440
      Max             =   0
      Min             =   -100
      TabIndex        =   11
      Top             =   4680
      Width           =   3615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.HScrollBar hsbPlayTime 
      Height          =   255
      Left            =   4440
      Max             =   100
      TabIndex        =   12
      Top             =   5760
      Width           =   3615
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "P&lay"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.DirListBox dirDirectoryListBox 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.FileListBox filMP3List 
      Height          =   3600
      Left            =   120
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   2640
      Width           =   3975
   End
   Begin VB.DriveListBox drvDriveSelect 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblPlaying 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4440
      TabIndex        =   17
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblPlayTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4440
      TabIndex        =   16
      Top             =   5280
      Width           =   1005
   End
   Begin VB.Label lblVolume 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4440
      TabIndex        =   15
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Label lblBalance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4440
      TabIndex        =   14
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label lblSpeed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
      Width           =   1005
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub cmdAbout_Click()
    Load frmAbout
    frmAbout.Show
    
End Sub

Private Sub cmdExit_Click()
    Call UnloadMP3Objects
    End
End Sub

Private Sub cmdOptions_Click()
    Load frmOptions
    frmOptions.Show
    
End Sub

Private Sub cmdPause_Click()
    Call PauseMP3
End Sub

Private Sub cmdPlay_Click()
    Call LoadMP3
    Call PlayMP3
End Sub

Private Sub cmdStop_Click()
    Call StopMP3
End Sub

Private Sub dirDirectoryListBox_Change()
    filMP3List.Path = dirDirectoryListBox.Path
End Sub

Private Sub drvDriveSelect_Change()
    dirDirectoryListBox.Path = drvDriveSelect.Drive
End Sub

Private Sub filMP3List_Click()
    Call LoadMP3
    Call PlayMP3
End Sub

Private Sub Form_Load()
        
    lblSpeed.Visible = False
    lblBalance.Visible = False
    hsbSpeed.Visible = False
    hsbBalance.Visible = False
    
    
    frmMain.Caption = App.Title
    
    If (mdlMain.PathMP3() = True) Then
        drvDriveSelect.Drive = strMP3Path
        dirDirectoryListBox.Path = strMP3Path
        
    End If
      
    
    filMP3List.Path = dirDirectoryListBox.Path
    
    frmMain.lblPlayTime.Caption = "PlayTime: 00:00"
    
    
    hsbSpeed.Value = 100
    
    Call hsbBalance_Change
    Call hsbVolume_Change
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnloadMP3Objects
        
End Sub

Private Sub hsbBalance_Change()
    'Abs Provents the "-" from showing up in the Caption
    lblBalance.Caption = "Balance: " & IIf(hsbBalance.Value < 0, Abs(hsbBalance.Value / 50) & "% Left", IIf(hsbBalance.Value > 0, (hsbBalance.Value / 50) & "% Right", "Absolute Center"))
    mdlMain.PlayBalance (hsbBalance.Value)
     
End Sub

Private Sub hsbPlayTime_Scroll()
    
    intMinutes = 0
    intSeconds = 0
       
    mdlMain.PlayTime (hsbPlayTime.Value)
End Sub

Private Sub hsbSpeed_Change()
    lblSpeed.Caption = "Speed: " & hsbSpeed.Value & "%"
    
    mdlMain.PlaySpeed (hsbSpeed.Value / 100)
    
End Sub

Private Sub hsbVolume_Change()
    lblVolume.Caption = "Volume: " & 100 + hsbVolume.Value & "%"
    mdlMain.PlayVolume (hsbVolume.Value)
        
End Sub

Private Sub tmrTimer_Timer()
    Call PlayTimeUpdate
    
End Sub
