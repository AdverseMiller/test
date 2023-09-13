Dim objShell, strHomeFolder, strLink
Set objShell = CreateObject("WScript.Shell")



' Get the home directory of the current user
strHomeFolder = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

' Get the first command line argument (the link)
If WScript.Arguments.Count > 0 Then
    strLink = WScript.Arguments(0)
Else
    WScript.Echo "No link provided. Exiting."
    WScript.Quit
End If

' Run the nircmd.exe commands
objShell.Run strHomeFolder & "\nircmd.exe mutesysvolume 0", 0, False
objShell.Run strHomeFolder & "\nircmd.exe setsysvolume 65535", 0, False


Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DisplayConfiguration")


' Run Chrome minimized with the given link and autoplay policy
objShell.Run "cmd.exe /c start /min chrome.exe " & strLink & " --autoplay-policy=no-user-gesture-required --window-position=" & objItem.PelsWidth-80 & "," & (objItem.PelsHeight * -1)+45, vbHide
Set objShell = Nothing
