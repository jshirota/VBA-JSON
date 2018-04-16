Attribute VB_Name = "MshtaModule"
Option Explicit

' source https://stackoverflow.com/questions/9725882/getting-scriptcontrol-to-work-with-excel-2010-x64

Function CreateObjectx86(Optional sProgID)

    Static oWnd As Object
    Dim bRunning As Boolean

    #If Win64 Then
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        Select Case True
            Case IsMissing(sProgID)
                If bRunning Then oWnd.Lost = False
                Exit Function
            Case IsEmpty(sProgID)
                If bRunning Then oWnd.Close
                Exit Function
            Case Not bRunning
                Set oWnd = CreateWindow()
                oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID) End Function", "VBScript"
                oWnd.execScript "var Lost, App;": Set oWnd.App = Application
                oWnd.execScript "Sub Check(): On Error Resume Next: Lost = True: App.Run(""CreateObjectx86""): If Lost And (Err.Number = 1004 Or Err.Number = 0) Then close: End If End Sub", "VBScript"
                oWnd.execScript "setInterval('Check();', 500);"
        End Select
        Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
    #Else
        Set CreateObjectx86 = CreateObject(sProgID)
    #End If

End Function

Function CreateWindow()

    ' source http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
    Dim sSignature, oShellWnd, oProc

    On Error Resume Next
    sSignature = Left(CreateObject("Scriptlet.TypeLib").GUID, 38)
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
            Set CreateWindow = oShellWnd.GetProperty(sSignature)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        Next
    Loop

End Function
