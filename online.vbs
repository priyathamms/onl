Dim obj

set objs = WScript.CreateObject("WScript.shell")
i=0

Do While i = 0
obj = objs.sendkeys("{NUMLOCK}{NUMLOCK}")
WScript.sleep(6000)
Loop
