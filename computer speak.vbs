Dim userinput
userinput = inputbox("Tell text to speak")
set sapi = wscript.createobject("SAPI.Spvoice")
Sapi.speak userinput


