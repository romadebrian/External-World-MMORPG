Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    ' Check if the config file exists.
    ' Don't randomly exit, at least notify the user.
    If Not FileExist(App.Path & "\Data Files\UpdateConfig.ini") Then
        MsgBox "The configuration file appears to be missing." & vbNewLine & "Please check if it exists or re-install the application.", vbCritical, "Error"
        DestroyUpdater
    End If
    
    ' Load the values we need into memory.
    UpdateURL = GetVar(App.Path & "\Data Files\UpdateConfig.ini", "UPDATER", "UpdateURL")
    NewsURL = GetVar(App.Path & "\Data Files\UpdateConfig.ini", "UPDATER", "NewsURL")
    GameWebsite = GetVar(App.Path & "\Data Files\UpdateConfig.ini", "UPDATER", "GameWebsite")
    GameName = GetVar(App.Path & "\Data Files\UpdateConfig.ini", "OPTIONS", "Game_Name")
    ClientName = Trim(GetVar(App.Path & "\Data Files\UpdateConfig.ini", "UPDATER", "ClientName"))
    
    ' Set the base value for a single %.
    ProgressP = 4.25 ' frmMain.picprogress.Width / 100
    
    ' Load the main form
    Load frmMain
End Sub

Public Sub DestroyUpdater()
    ' Delete all temporary doodles.
    If FileExist(App.Path & "\tmpUpdate.ini") Then Kill App.Path & "\tmpUpdate.ini"
    
    ' End the program.
    frmMain.inetDownload.Cancel
    Unload frmMain
    End
End Sub

Public Sub ChangeStatus(ByVal NewStatus As String)
    frmMain.lblstatus.Caption = NewStatus
End Sub

Public Sub CheckVersion()
Dim Filename As String
    
    On Error GoTo ErrorHandle
    
    ' Enable our timeout timer, so it doesn't endlessly keep
    ' trying to connect.
    frmMain.tmrTimeout.Enabled = True
    
    ' Change the status of our updater, and progress down to 0.
    ChangeStatus "Connecting to update server..."
    SetProgress 0
    
    ' Get the file which contains the info of updated files
    DownloadFile UpdateURL & "update.txt", App.Path & "\tmpUpdate.ini"
    
    ' Done with the download, update the progress and continue!
    SetProgress 50
    ChangeStatus "Retrieving version information."
    
    ' read the version count
    VersionCount = GetVar(App.Path & "\tmpUpdate.ini", "UPDATER", "Version")
    
    ' check if we've got a current client version saved
    If FileExist(App.Path & "\Data Files\Config.ini") Then
        CurVersion = Val(GetVar(App.Path & "\Data Files\Config.ini", "UPDATER", "Version"))
    Else
        CurVersion = 0
    End If
    
    SetProgress 100
    
    ' Disable it, we have progress!
    frmMain.tmrTimeout.Enabled = False
    
    ' are we up to date?
    If CurVersion < VersionCount Then
        UpToDate = 0
        ChangeStatus "Your client is outdated!"
        VersionsToGo = VersionCount - CurVersion
        PercentToGo = 100 / (VersionsToGo * 2)
        RunUpdates
    Else
        UpToDate = 1
        ChangeStatus "Your client is up to date!"
        
        ' Load a GUI image, if it does not exist.. Exit out of the program.
        Form_LoadPicture (App.Path & "\Data Files\graphics\gui\updater\launch.jpg")
        
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox "An unexpected error has occured: " & Err.Description & vbNewLine & vbNewLine & "It is likely that your configuration is incorrect.", vbCritical
    Err.Clear
    DestroyUpdater
End Sub

Public Sub RunUpdates()
Dim Filename As String
Dim i As Long
Dim UpdateID As Long
Dim CurProgress As Long
    
    On Error GoTo ErrorHandle
    
    If CurVersion = 0 Then CurVersion = 1
    UpdateID = 0
    CurProgress = 0
    ' loop around, download and unrar each update
    For i = CurVersion To VersionCount
        ' Increase Update ID by 1
        UpdateID = UpdateID + 1
        
        ' let them know we're actually doing something..
        CurProgress = CurProgress + PercentToGo
        ChangeStatus "Downloading update " & Str(UpdateID) & "/" & Str(VersionsToGo) & "..."
        SetProgress CurProgress
        
        ' Download time!
        Filename = "version" & Trim(Str(i)) & ".rar"
        DownloadFile UpdateURL & "/" & Filename, App.Path & "\" & Filename
        
        ' Done downloading? Awesome.. Time to change the status
        CurProgress = CurProgress + PercentToGo
        ChangeStatus "Installing update " & Str(UpdateID) & "/" & Str(VersionsToGo) & "..."
        SetProgress CurProgress
        
        ' Extract date from the update file.
        RARExecute OP_EXTRACT, Filename
        
        ' Delete the update file.
        Kill App.Path & "\" & Filename
    Next
    
    ' Update the version of the client.
    PutVar App.Path & "\Data Files\Config.ini", "UPDATER", "Version", Str(VersionCount)
    
    ' Done? Niiiice..
    SetProgress 100 'Just to be sure, sometimes it misses ~1% due to the lack of decimals.
    ChangeStatus "Your client is up to date!"
    
     ' Load the Launch Backdrop.
    Form_LoadPicture (App.Path & "\Data Files\graphics\gui\updater\launch.jpg")
    
    ' Set the Update variable to 1, to prevent running this sub again.
    UpToDate = 1
    
    Exit Sub
    
ErrorHandle:
    MsgBox "An error has occured while extracting the update(s): " & Err.Description & vbNewLine & vbNewLine & "Please relay this message to an administrator.", vbCritical
    Err.Clear
    DestroyUpdater
End Sub

Public Sub SetProgress(ByVal Percent As Long)
    If Percent = 0 Then
        frmMain.picprogress.Width = 0
    Else
        frmMain.picprogress.Width = (ProgressP * Percent)
    End If
End Sub

Private Sub DownloadFile(ByVal URL As String, ByVal Filename As String)
    Dim fileBytes() As Byte
    Dim fileNum As Integer
    
    On Error GoTo DownloadError
    
    ' download data to byte array
    fileBytes() = frmMain.inetDownload.OpenURL(URL, icByteArray)
    
    fileNum = FreeFile
    Open Filename For Binary Access Write As #fileNum
        ' dump the byte array as binary
        Put #fileNum, , fileBytes()
    Close #fileNum
    
    Exit Sub
    
DownloadError:
    MsgBox Err.Description
End Sub

Public Sub Form_LoadPicture(ByVal Filename As String)
    If FileExist(Filename) Then
            frmMain.Picture = LoadPicture(Filename)
        Else
            DestroyUpdater
        End If
End Sub
