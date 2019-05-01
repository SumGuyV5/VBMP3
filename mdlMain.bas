Attribute VB_Name = "mdlMain"
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

Public strMP3Path As String                'The Folder to load up first

Private IBAudio  As IBasicAudio         'Basic Audio Object
Private IMEvent As IMediaEvent          'MediaEvent Object
Private IMControl As IMediaControl      'MediaControl Object
Private IMPosition As IMediaPosition    'MediaPosition Object

Dim strTotalPlayTime As String

Dim intPosition As Integer
Public intMinutes As Integer
Public intSeconds As Integer

Dim intDuration As Integer
Dim intTotalMinutes As Integer
Dim intTotalSeconds As Integer

Private blnRunningMP3 As Boolean        'Is the MP3 Running?
Private blnPauseMP3 As Boolean          'Has the MP3 Been Paused?
Private blnLoadedMP3 As Boolean         'Has the MP3 Been Load?

Public Function LoadMP3()
On Error GoTo BailOut:
             
        frmMain.Caption = App.Title
             
    'Setup a filter graph for the file
        Set IMControl = New FilgraphManager
        IMControl.RenderFile (frmMain.dirDirectoryListBox.Path & "\" & frmMain.filMP3List.FileName)
                        
    'Setup the basic audio object
        Set IBAudio = IMControl
        IBAudio.Volume = frmMain.hsbVolume.Value * 40
        IBAudio.Balance = frmMain.hsbBalance.Value
    
    'Setup the media event and intPosition objects
        Set IMEvent = IMControl
        Set IMPosition = IMControl
        'IMPosition.Rate = PlaySpeed(frmMain.hsbSpeed.Value )
        'IMPosition.Rate = 1
        IMPosition.Rate = frmMain.hsbSpeed.Value / 100
        IMPosition.CurrentPosition = 0
        'IMPosition.CurrentPosition = frmMain.hsbPlayTime.Value
        
        intDuration = IMPosition.Duration
        
        'Sets the Varables for finding out the Current Position Time back to 0
        intPosition = 0
        intMinutes = 0
        intSeconds = 0
        
        'Debug.Print IMPosition.PrerollTime
        
        
        frmMain.hsbPlayTime.Max = intDuration + 1
        
        blnLoadedMP3 = True
        
        frmMain.Caption = frmMain.Caption & " - " & frmMain.filMP3List.FileName
        
        'New Code April 8, 2003
        strTotalPlayTime = TotalTime
        'End of New Code
            
           
    Exit Function
BailOut:
    blnLoadedMP3 = False
    MsgBox "Unable to Load the Mp3!!", vbCritical, "Error"
    Debug.Print "ERROR: mdlMain.LoadMP3()"
    Debug.Print Err.Number, Err.Description
End Function

Public Function PlayMP3()
On Error GoTo BailOut:

    IMControl.Run   'Runs the MP3
       
    frmMain.lblPlaying.Caption = frmMain.filMP3List.FileName
    
   
    blnRunningMP3 = True
    Exit Function
BailOut:
    MsgBox "Unable to Playing the Mp3!!", vbCritical, "Error"
    Debug.Print "ERROR: mdlMain.Play()"
    Debug.Print Err.Number, Err.Description
End Function

Public Function StopMP3()
On Error GoTo BailOut:
    IMControl.Stop
    IMPosition.CurrentPosition = 0 'set it back to the beginning
    intMinutes = 0
    intSeconds = 0
    blnRunningMP3 = False
    
    Exit Function
BailOut:
    MsgBox "Unable to Stop playing the Mp3!!", vbCritical, "Error"
    Debug.Print "ERROR: mdlMain.StopMP3()"
    Debug.Print Err.Number, Err.Description
End Function

Public Function PauseMP3()
On Error GoTo BailOut:
    If blnRunningMP3 = True Then
        If blnPauseMP3 = True Then
            IMControl.Run
            blnPauseMP3 = False
        Else
            IMControl.Pause
            blnPauseMP3 = True
        End If
    End If
    
    Exit Function
BailOut:
    MsgBox "Unable to Pause the Mp3!!", vbCritical, "Error"
    Debug.Print "ERROR: mdlMain.PauseMP3()"
    Debug.Print Err.Number, Err.Description
End Function

'Changes the Speed that MP3 Playes at.
Public Function PlaySpeed(Speed As Single)
On Error GoTo BailOut:

    If ObjPtr(IMPosition) > 0 Then
        IMPosition.Rate = Speed
    End If
    
    Exit Function
BailOut:
    MsgBox "Unable to Change the Speed on the Mp3!!", vbCritical, "Error"
    Debug.Print "ERROR: mdlMain.PlaySpeed()"
    Debug.Print Err.Number, Err.Description

End Function

'Changes the PlayVolume of the MP3.
Public Function PlayVolume(Volume As Long)
On Error GoTo BailOut:
    If ObjPtr(IMControl) > 0 Then
        IBAudio.Volume = Volume * 40 ' makes it in the 0 to -4000 range (-4000 is almost silent)
    End If
    
    Exit Function
BailOut:
    MsgBox "Unable to Change the Volume on the Mp3!!", vbCritical, "Error"
    Debug.Print "ERROR: mdlMain.PlayVolume()"
    Debug.Print Err.Number, Err.Description
End Function

'Balances the Volume for the left and right speaker.
Public Function PlayBalance(Balance As Long)
On Error GoTo BailOut:

    If ObjPtr(IMControl) > 0 Then
        IBAudio.Balance = Balance
    End If
    
    Exit Function
BailOut:
    MsgBox "Unable to Change the Balance on the Mp3!!", vbCritical, "Error"
    Debug.Print "ERROR: mdlMain.PlayBalance()"
    Debug.Print Err.Number, Err.Description
End Function

'Sets the MP3 to a new Position
'This function will change the currentposition of the MP3 relative to hsbPlayTime.
Public Function PlayTime(Time As Integer)
On Error GoTo BailOut:
    
    If ObjPtr(IMPosition) > 0 Then
        IMPosition.CurrentPosition = Time
    End If
        
    Exit Function
BailOut:
    MsgBox "Unable to Change the intPosition on the Mp3!!", vbCritical, "Error"
    Debug.Print "ERROR: mdlMain.Playtime()"
    Debug.Print Err.Number, Err.Description
End Function

Public Function UnloadMP3Objects()
    
    'set all objects to Nothing
    Set IBAudio = Nothing
    Set IMEvent = Nothing
    Set IMControl = Nothing
    Set IMPosition = Nothing
    
End Function

'Used to find out how far you are in to the MP3
Public Function PlayTimeUpdate()
                
    If blnLoadedMP3 = True Then
        
        'Gets the CurrentPosition in seconds
        intPosition = IMPosition.CurrentPosition
        
        'We have to find out if intPosition is 60 seconds more then the Current time
        'To do this we take the Current Minutes (intMinutes) add one to it and Divide it by 60
        'to get the total number of seconds that are need to increment intMinutes to the next level
        'Note: I Could have also used this but did not think of it at the time.
        'If intSeconds >= 60 Then
        '   intSeconds = 0
        '    intMinutes = intMinutes + 1
        'End If
        If intPosition > ((intMinutes + 1) * 60) Then
            intMinutes = intMinutes + 1
        End If
        
        'Find out how many seconds we have
        intSeconds = intPosition - (intMinutes * 60)
                 
        
        'New Faster Code April 8, 2003
        'This if statment is for formating perpase
        'Here we find out if the inMinutes is Single or Double digits
        If (intMinutes < 10) Then
            'Same thing for intSeconds
            If (intSeconds < 10) Then
                frmMain.lblPlayTime.Caption = "PlayTime: 0" & intMinutes & ":0" & intSeconds & " / " & strTotalPlayTime
            Else
                frmMain.lblPlayTime.Caption = "PlayTime: 0" & intMinutes & ":" & intSeconds & " / " & strTotalPlayTime
                
            End If
        Else
            If (intSeconds < 10) Then
                frmMain.lblPlayTime.Caption = "PlayTime: " & intMinutes & ":0" & intSeconds & " / " & strTotalPlayTime
            Else
                frmMain.lblPlayTime.Caption = "PlayTime: " & intMinutes & ":" & intSeconds & " / " & strTotalPlayTime
            End If
        End If
        'End New Faster Code
              
        'The Old Code Would Call TotalTime function Ever time the Current time was printed.
        'In the New way when the MP3 is load it calls the TotalTime to get the total time and stores
        'it in from strTotalPlayTime now when ever the Playtime is update it just puts strTotalPlayTime on end
        
        'Old Slow Code
        'If (intMinutes < 10) Then
        '    'Same thing for intSeconds
        '    If (intSeconds < 10) Then
        '        frmMain.lblPlayTime.Caption = "PlayTime: 0" & intMinutes & ":0" & intSeconds & " / " & TotalTime
        '    Else
        '        frmMain.lblPlayTime.Caption = "PlayTime: 0" & intMinutes & ":" & intSeconds & " / " & TotalTime
        '
        '    End If
        'Else
        '    If (intSeconds < 10) Then
        '        frmMain.lblPlayTime.Caption = "PlayTime: " & intMinutes & ":0" & intSeconds & " / " & TotalTime
        '    Else
        '        frmMain.lblPlayTime.Caption = "PlayTime: " & intMinutes & ":" & intSeconds & " / " & TotalTime
        '    End If
        'End If
        'End Old Slow Code
        
        
        frmMain.hsbPlayTime.Value = IMPosition.CurrentPosition
    End If
End Function

Private Function TotalTime()
       
    intTotalMinutes = intDuration / 60                     'We Need to get the total time in Minutes
    
    'This line is to see if intTotalMinutes was rounded up as it is in integer
        'We can't compare Minutes to Seconds so we multiply intTotalMinutes by 60 and the check
    If ((intTotalMinutes * 60) >= intDuration) Then
        intTotalMinutes = intTotalMinutes - 1                   'We subtract a Minute do to Integers rounding up of the number
    End If
    
    'We Get intTotalSeconds by take the intDuration
    'Here we take the intTotalMinutes and agen change to seconds by multiply
    intTotalSeconds = intDuration - (intTotalMinutes * 60)
    
    'Here we find out if the inTotalMinutes is Single or Double digits
    If (intTotalMinutes < 10) Then
        'Same thing for intTotalSeconds
        If (intTotalSeconds < 10) Then
            TotalTime = "0" & intTotalMinutes & ":0" & intTotalSeconds
        Else
            TotalTime = "0" & intTotalMinutes & ":" & intTotalSeconds
        End If
    Else
        If (intTotalSeconds < 10) Then
            TotalTime = intTotalMinutes & ":0" & intTotalSeconds
        Else
            TotalTime = intTotalMinutes & ":" & intTotalSeconds
        End If
    End If
    
    
End Function

Public Function PathMP3()
On Error GoTo OpenErrHandler
    Open App.Path & "/Settings.dat" For Input As #1
        Input #1, strMP3Path
    Close #1
    PathMP3 = True
    Exit Function
OpenErrHandler:
    PathMP3 = False
End Function
