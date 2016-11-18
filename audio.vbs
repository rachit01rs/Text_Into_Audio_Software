Dim message, sapi
 message = InputBox("A Best Text to Audio converter"+vbcrlf+"From - www.rabins.com","Text to Audio converter")
 Set sapi = CreateObject("sapi.spvoice")
 sapi.Speak message