Dim message,sapi
message=InputBox("TTS")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message