inpt = inputbox("Geben Sie bitte die WLAN SSID")
set oShell = WScript.CreateObject("WScript.Shell")
oShell.run("netsh wlan connect name=" & inpt)