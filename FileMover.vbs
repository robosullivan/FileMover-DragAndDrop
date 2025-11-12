' === FileMover : Drag-and-drop === - Move files with undo, logging, balloon tips, and smart hiding of config files
' Written to simplify moving files and keep track of actions for undo and auditing.

Option Explicit

' Initialize required objects: FileSystem, Shell, Network
Dim fso, shell, net, objArgs
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set net = CreateObject("WScript.Network")
Set objArgs = WScript.Arguments

' Define paths for config and logs based on script location and name
Dim scriptPath, scriptFolder, scriptBase
scriptPath = WScript.ScriptFullName
scriptFolder = fso.GetParentFolderName(scriptPath)
scriptBase = fso.GetBaseName(WScript.ScriptName)

Dim configPath, logFolder, logFile, logDate
configPath = scriptFolder & "\" & scriptBase & "_destination_folder.txt"
logFolder = scriptFolder & "\logs\"
' === Prevent config file from being moved ===
If fso.FileExists(configPath) Then
    Dim actualConfigFolder
    actualConfigFolder = fso.GetParentFolderName(configPath)
    If LCase(actualConfigFolder) <> LCase(scriptFolder) Then
        MsgBox "The configuration file '" & configPath & "' has been moved from its original location." & vbCrLf & _
               "Please move it back to the script's folder: " & scriptFolder, vbCritical, "Configuration File Moved"
        WScript.Quit
    End If
End If

logDate = Year(Now) & "-" & Right("0" & Month(Now),2) & "-" & Right("0" & Day(Now),2)
logFile = logFolder & scriptBase & "_" & logDate & ".log"

Dim destFolder, totalMoved, renamedCount, folderSkippedCount
totalMoved = 0 : renamedCount = 0 : folderSkippedCount = 0

' === MAIN EXECUTION FLOW ===
' Handle cases when no files are dropped: first time setup, undo prompt, or exit
If objArgs.Count = 0 Then
    If Not fso.FileExists(configPath) Then
        ' First-time setup
        destFolder = GetOrCreateDestination(configPath)
        If Not fso.FolderExists(logFolder) Then fso.CreateFolder(logFolder)
        
        ' Write dummy log to prevent double prompt
        WriteDummyLog
        
        PromptHideFileOptions configPath, logFolder
        MsgBox "Setup complete. You can now drag files onto this script.", vbInformation, "FileMover Setup"
    ElseIf AskUndo() Then
        Call UndoLastMove()
    Else
        WScript.Echo "No files dropped. Exiting."
    End If
    WScript.Quit
End If

' Proceed with file move
destFolder = GetOrCreateDestination(configPath)

' Check if destination folder is writable before proceeding
If Not IsFolderWritable(destFolder) Then
    MsgBox "Cannot write to destination folder: " & destFolder, vbExclamation, "Permission Error"
    WScript.Quit
End If

' Create log folder if missing
If Not fso.FolderExists(logFolder) Then fso.CreateFolder(logFolder)

' Prompt about hiding files only if no previous logs exist (first use)
If Not HasPreviousLog(logFolder, scriptBase) Then
    PromptHideFileOptions configPath, logFolder
End If

' Loop through all dropped files and move them
Dim i
For i = 0 To objArgs.Count - 1
    Dim sourceFile : sourceFile = objArgs(i)
    If fso.FolderExists(sourceFile) Then
        folderSkippedCount = folderSkippedCount + 1
        WriteLog "SKIP", sourceFile, "-", "Skipped: folder"
    Else
        Dim fileName : fileName = fso.GetFileName(sourceFile)
        Dim uniqueName : uniqueName = GetUniqueFileName(destFolder, fileName)
        Dim destFile : destFile = destFolder & uniqueName
        On Error Resume Next
        fso.MoveFile sourceFile, destFile
        If Err.Number <> 0 Then
            WriteLog "MOVE", sourceFile, destFile, "ERROR: " & Err.Description
            Err.Clear
        Else
            WriteLog "MOVE", sourceFile, destFile, "SUCCESS"
            totalMoved = totalMoved + 1
            If uniqueName <> fileName Then renamedCount = renamedCount + 1
        End If
        On Error GoTo 0
    End If
Next

' Show balloon notification summarizing the file move operation
ShowBalloonSummary totalMoved, renamedCount, folderSkippedCount

' === FUNCTIONS ===

Function GetOrCreateDestination(configFile)
    ' Returns the destination folder path from config or asks user to select one
    Dim path
    If fso.FileExists(configFile) Then
        Dim tf : Set tf = fso.OpenTextFile(configFile, 1)
        path = Trim(tf.ReadLine)
        tf.Close
        If Not fso.FolderExists(path) Then
            path = PromptForFolder()
        End If
    Else
        path = PromptForFolder()
        If path <> "" Then
            Dim f : Set f = fso.CreateTextFile(configFile, True)
            f.WriteLine path
            f.Close
        End If
    End If
    If path = "" Then
        WScript.Echo "No destination selected. Exiting."
        WScript.Quit
    End If
    If Right(path,1) <> "\" Then path = path & "\"
    GetOrCreateDestination = path
End Function

Function PromptForFolder()
    Dim choice
    choice = MsgBox("Welcome to FileMover:DragAndDrop!" & vbCrLf & _
                    "Just drop files onto the FileMover icon and they'll be moved to your chosen folder." & vbCrLf & vbCrLf & _
                    "Lets set that folder now." & vbCrLf & vbCrLf & _
                    "Would you like to browse for the destination folder?" & vbCrLf & vbCrLf & _
                    "Yes = Browse for folder" & vbCrLf & _
                    "No = Type folder path manually", vbYesNoCancel + vbQuestion, "FileMover:DragAndDrop - Setup")

    If choice = vbYes Then
        PromptForFolder = BrowseForFolder("Select the destination folder:")
    ElseIf choice = vbNo Then
        Dim inputPath
        inputPath = InputBox("Enter the full path to the destination folder (UNC paths allowed):", "Type Folder Path")
        If inputPath <> "" Then
            If fso.FolderExists(inputPath) Then
                PromptForFolder = inputPath
            Else
                MsgBox "Unable to locate folder path:" & vbCrLf & inputPath, vbExclamation, "Invalid Path"
                PromptForFolder = ""
            End If
        Else
            PromptForFolder = ""
        End If
    Else
        PromptForFolder = ""
    End If
End Function

Function BrowseForFolder(prompt)
    Dim shellApp, folder, startLocation
    Set shellApp = CreateObject("Shell.Application")

    ' Determine starting point: one folder up if on network, else "This PC"
    If Left(scriptFolder, 2) = "\\" Then
        startLocation = fso.GetParentFolderName(scriptFolder)
    Else
        startLocation = &H0011  ' CSIDL_DRIVES ("This PC")
    End If

    ' Show folder picker WITHOUT type box, with modern UI
    Const OPTIONS = &H0001  ' BIF_NEWDIALOGSTYLE only
    Set folder = shellApp.BrowseForFolder(0, prompt, OPTIONS, startLocation)

    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path & "\"
    Else
        BrowseForFolder = ""
    End If
End Function




Function IsFolderWritable(folderPath)
    ' Checks if folder is writable by trying to create and delete a temp file
    On Error Resume Next
    Dim testFile, ts
    testFile = folderPath & "write_test_" & Replace(Replace(CreateObject("Scriptlet.TypeLib").GUID,"{",""),"}","") & ".tmp"
    Set ts = fso.CreateTextFile(testFile, True)
    ts.WriteLine "test" : ts.Close
    fso.DeleteFile(testFile)
    IsFolderWritable = (Err.Number = 0)
    Err.Clear : On Error GoTo 0
End Function

Function GetUniqueFileName(folderPath, fileName)
    ' Returns a unique file name by appending (1), (2), etc. if file exists
    Dim base, ext, newName, i
    base = fso.GetBaseName(fileName)
    ext = fso.GetExtensionName(fileName)
    If ext <> "" Then ext = "." & ext
    newName = fileName : i = 1
    While fso.FileExists(folderPath & newName)
        newName = base & " (" & i & ")" & ext
        i = i + 1
    Wend
    GetUniqueFileName = newName
End Function

Sub WriteLog(action, src, dest, status)
    ' Writes a log entry to the daily log file with timestamp and user info
    Dim logEntry, f, user
    user = net.UserName
    logEntry = "[" & Now & "] | USER: " & user & " | " & action & " | " & src & " | " & dest & " | " & status
    Set f = fso.OpenTextFile(logFile, 8, True)
    f.WriteLine logEntry
    f.Close
End Sub

' === Write a dummy log entry for first run to avoid double prompt ===
Sub WriteDummyLog()
    Dim f, dummyEntry, user
    user = net.UserName
    dummyEntry = "[" & Now & "] | USER: " & user & " | DUMMY | Initialization log entry | - | SUCCESS"
    Set f = fso.OpenTextFile(logFile, 8, True)
    f.WriteLine dummyEntry
    f.Close
End Sub

Function HasPreviousLog(folderPath, base)
    ' Checks if any previous log files exist in log folder matching base name
    On Error Resume Next
    If Not fso.FolderExists(folderPath) Then Exit Function
    Dim file
    For Each file In fso.GetFolder(folderPath).Files
        If LCase(Left(file.Name, Len(base))) = LCase(base) Then
            HasPreviousLog = True
            Exit Function
        End If
    Next
    HasPreviousLog = False
End Function

' Prompts user to optionally hide config file and log folder
Sub PromptHideFileOptions(cfgPath, logPath)
    Dim msg, choice
    msg = "Setup is almost complete." & vbCrLf & vbCrLf & _
          "Would you like to hide BOTH the configuration file and the log folder?" & vbCrLf & vbCrLf & _
          "Hiding them keeps them less visible but they can be unhidden manually later."

    choice = MsgBox(msg, vbYesNo + vbQuestion, "Hide Config and Logs?")

    If choice = vbYes Then
        If fso.FileExists(cfgPath) Then fso.GetFile(cfgPath).Attributes = fso.GetFile(cfgPath).Attributes Or 2
        If fso.FolderExists(logPath) Then fso.GetFolder(logPath).Attributes = fso.GetFolder(logPath).Attributes Or 2
    End If
    ' If No, do nothing (leave visible)
End Sub



Sub ShowBalloonSummary(moved, renamed, skipped)
    Dim summary, singleFileName

    If moved = 1 And objArgs.Count = 1 Then
        ' Only one file moved, show first 30 characters of the file name
        singleFileName = fso.GetFileName(objArgs(0))
        If Len(singleFileName) > 30 Then
            singleFileName = Left(singleFileName, 30) & "..."
        End If
        summary = "Moved file: " & singleFileName
        If renamed > 0 Then summary = summary & " (renamed)"
        If skipped > 0 Then summary = summary & ". " & skipped & " folder(s) skipped."
    Else
        ' Multiple files or other cases: keep original summary format
        summary = moved & " file"
        If moved <> 1 Then summary = summary & "s"
        summary = summary & " moved."
        If renamed > 0 Then summary = summary & " " & renamed & " renamed."
        If skipped > 0 Then summary = summary & " " & skipped & " folder(s) skipped."
    End If

    summary = Replace(summary, "'", "''")
    
    Dim ps
    ps = "powershell -Command ""Add-Type -AssemblyName System.Windows.Forms;" & _
         "$n = New-Object Windows.Forms.NotifyIcon;" & _
         "$n.Icon = [System.Drawing.SystemIcons]::Information;" & _
         "$n.BalloonTipTitle = '" & scriptBase & "';" & _
         "$n.BalloonTipText = '" & summary & "';" & _
         "$n.Visible = $true;" & _
         "$n.ShowBalloonTip(3000);" & _
         "Start-Sleep -Seconds 4;" & _
         "$n.Dispose();"""
    shell.Run ps, 0, False
End Sub


Function AskUndo()
    ' Asks user if they want to undo the last move if no files were dropped
    AskUndo = (MsgBox("No files dropped." & vbCrLf & _
                      "Do you want to undo the last move?", _
                      vbYesNo + vbQuestion, "Undo Last Move?") = vbYes)
End Function

Sub UndoLastMove()
    ' Attempts to undo the last file move based on today's log entries
    If Not fso.FileExists(logFile) Then
        MsgBox "No log found for today: " & logFile, vbExclamation, "Undo Failed"
        Exit Sub
    End If
    Dim lines, line, f, fields, src, dest
    Set f = fso.OpenTextFile(logFile, 1)
    lines = Split(f.ReadAll, vbCrLf)
    f.Close
    Dim i
    For i = UBound(lines) To 0 Step -1
        line = Trim(lines(i))
        If line <> "" Then
            fields = Split(line, "|")
            If UBound(fields) >= 5 Then
                If Trim(fields(2)) = "MOVE" And Trim(fields(5)) = "SUCCESS" Then
                    src = Trim(fields(3))
                    dest = Trim(fields(4))
                    If fso.FileExists(dest) Then
                        On Error Resume Next
                        fso.MoveFile dest, src
                        If Err.Number = 0 Then
                            WriteLog "UNDO", dest, src, "SUCCESS"
                            MsgBox "Undo successful: " & fso.GetFileName(dest), vbInformation, "Undo"
                        Else
                            WriteLog "UNDO", dest, src, "ERROR: " & Err.Description
                            MsgBox "Undo failed: " & Err.Description, vbExclamation, "Undo"
                        End If
                        Err.Clear
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    MsgBox "No eligible move found to undo.", vbExclamation, "Undo"
End Sub
