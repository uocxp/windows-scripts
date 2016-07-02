inpt = inputbox("Give please the wlan ssid")
set oShell = WScript.CreateObject("WScript.Shell")
oShell.run("netsh wlan connect name=" & inpt)