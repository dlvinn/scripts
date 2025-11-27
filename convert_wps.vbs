' WPS Office ODT to DOCX Converter
' This script uses WPS Office COM automation to convert ODT files to DOCX

Option Explicit

Dim objFSO, objShell, objFolder, objFile
Dim wpsApp, wpsDoc
Dim strFolder, strInputFile, strOutputFile
Dim intCount, intSuccess, intFailed

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Get current directory
strFolder = objShell.CurrentDirectory

WScript.Echo "========================================"
WScript.Echo "WPS Office ODT to DOCX Converter"
WScript.Echo "========================================"
WScript.Echo ""
WScript.Echo "Scanning folder: " & strFolder
WScript.Echo ""

' Initialize counters
intCount = 0
intSuccess = 0
intFailed = 0

' Try to create WPS application object
On Error Resume Next
Set wpsApp = CreateObject("WPS.Application")

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot create WPS Application object"
    WScript.Echo "Make sure WPS Office is installed properly"
    WScript.Echo ""
    WScript.Echo "Press any key to exit..."
    WScript.StdIn.ReadLine
    WScript.Quit 1
End If

On Error GoTo 0

' Make WPS invisible
wpsApp.Visible = False

' Scan for ODT files recursively
ScanFolder objFSO.GetFolder(strFolder)

' Clean up
wpsApp.Quit
Set wpsApp = Nothing

' Show summary
WScript.Echo ""
WScript.Echo "========================================"
WScript.Echo "CONVERSION SUMMARY"
WScript.Echo "========================================"

If intCount = 0 Then
    WScript.Echo "No .odt files found"
Else
    WScript.Echo "Total files found:      " & intCount
    WScript.Echo "Successfully converted: " & intSuccess
    WScript.Echo "Failed:                 " & intFailed
End If

WScript.Echo "========================================"
WScript.Echo ""
WScript.Echo "Press Enter to exit..."
WScript.StdIn.ReadLine

WScript.Quit 0

' Recursive folder scanning
Sub ScanFolder(folder)
    Dim file, subfolder

    ' Process files in current folder
    For Each file In folder.Files
        If LCase(objFSO.GetExtensionName(file.Path)) = "odt" Then
            intCount = intCount + 1
            ConvertFile file.Path
        End If
    Next

    ' Process subfolders
    For Each subfolder In folder.SubFolders
        ScanFolder subfolder
    Next
End Sub

' Convert single file
Sub ConvertFile(inputPath)
    Dim outputPath, fileName

    fileName = objFSO.GetFileName(inputPath)
    outputPath = objFSO.GetParentFolderName(inputPath) & "\" & _
                 objFSO.GetBaseName(inputPath) & ".docx"

    WScript.Echo "[" & intCount & "] Converting: " & fileName
    WScript.Echo "    Location: " & objFSO.GetParentFolderName(inputPath)

    On Error Resume Next

    ' Open ODT file
    Set wpsDoc = wpsApp.Documents.Open(inputPath)

    If Err.Number <> 0 Then
        WScript.Echo "    Status: FAILED (cannot open file)"
        WScript.Echo "    Error: " & Err.Description
        intFailed = intFailed + 1
        Err.Clear
        Exit Sub
    End If

    ' Save as DOCX
    ' wdFormatXMLDocument = 12 (DOCX format)
    wpsDoc.SaveAs2 outputPath, 12

    If Err.Number <> 0 Then
        WScript.Echo "    Status: FAILED (cannot save file)"
        WScript.Echo "    Error: " & Err.Description
        intFailed = intFailed + 1
        Err.Clear
    Else
        WScript.Echo "    Status: SUCCESS"
        intSuccess = intSuccess + 1
    End If

    ' Close document
    wpsDoc.Close False
    Set wpsDoc = Nothing

    On Error GoTo 0

    WScript.Echo ""
End Sub
