' Check for all available Office applications
On Error Resume Next

WScript.Echo "Checking for Office applications..."
WScript.Echo ""

' Check WPS Office
Set wpsApp = CreateObject("WPS.Application")
If Err.Number = 0 Then
    WScript.Echo "[FOUND] WPS Office (WPS.Application)"
    wpsApp.Quit
Else
    WScript.Echo "[NOT FOUND] WPS Office (WPS.Application)"
End If
Err.Clear

' Check Kingsoft WPS
Set ksoApp = CreateObject("KSO.Application")
If Err.Number = 0 Then
    WScript.Echo "[FOUND] Kingsoft Office (KSO.Application)"
    ksoApp.Quit
Else
    WScript.Echo "[NOT FOUND] Kingsoft Office (KSO.Application)"
End If
Err.Clear

' Check Microsoft Word
Set wordApp = CreateObject("Word.Application")
If Err.Number = 0 Then
    WScript.Echo "[FOUND] Microsoft Word (Word.Application)"
    wordApp.Quit
Else
    WScript.Echo "[NOT FOUND] Microsoft Word (Word.Application)"
End If
Err.Clear

' Check LibreOffice
Set officeApp = CreateObject("com.sun.star.ServiceManager")
If Err.Number = 0 Then
    WScript.Echo "[FOUND] LibreOffice (com.sun.star.ServiceManager)"
Else
    WScript.Echo "[NOT FOUND] LibreOffice (com.sun.star.ServiceManager)"
End If

WScript.Echo ""
WScript.Echo "Press Enter to exit..."
WScript.StdIn.ReadLine
