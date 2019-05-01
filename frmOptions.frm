VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CommandButton cmdSetBackGround 
      Caption         =   "Set Back &Ground"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetDefaltFolder 
      Caption         =   "Set &Defalt Folder"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdResetControls 
      Caption         =   "&Reset Controls"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox chkAdvanceControls 
      Caption         =   "&Advance Controls"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmOptions"
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

Private Sub cmd_Click()
    frmOptions.Hide
    Unload frmOptions
    
End Sub

Private Sub cmdResetControls_Click()
    frmMain.hsbBalance.Value = 0
    frmMain.hsbSpeed.Value = 100
    frmMain.hsbVolume.Value = 0
    
End Sub

Private Sub cmdSetBackGround_Click()
    MsgBox "This Option is not available at this time!!", vbCritical, "Error"
End Sub

Private Sub cmdSetDefaltFolder_Click()
    Load frmDefaultDir
    frmDefaultDir.Show
End Sub

Private Sub Form_Load()
    If frmMain.lblSpeed.Visible = True Then
        chkAdvanceControls.Value = 1
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If chkAdvanceControls.Value = 1 Then
        frmMain.lblSpeed.Visible = True
        frmMain.lblBalance.Visible = True
        frmMain.hsbSpeed.Visible = True
        frmMain.hsbBalance.Visible = True
    Else
        frmMain.lblSpeed.Visible = False
        frmMain.lblBalance.Visible = False
        frmMain.hsbSpeed.Visible = False
        frmMain.hsbBalance.Visible = False
    End If
End Sub
