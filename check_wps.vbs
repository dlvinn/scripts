' Quick check if WPS Office is available via COM
On Error Resume Next
Dim wpsApp
Set wpsApp = CreateObject("WPS.Application")
If Err.Number = 0 Then
    wpsApp.Quit
    WScript.Quit 0
Else
    WScript.Quit 1
End If
