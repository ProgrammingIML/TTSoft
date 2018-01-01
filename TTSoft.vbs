Dim retval
    retval = MsgBox("Hello! TTSoft V1.7",1 + vbDefaultButton1,"TTSoft")
	
	Dim message, sapi
	Set sapi=CreateObject("sapi.spvoice")
	sapi.Rate = 1
	sapi.Volume = 100
	do
	message=InputBox("Enter Text.","TTSoft","Insert text here.")

	if message <> "" then sapi.Speak message
	loop Until message=""
	
Dim ExitVal
	ExitVal = Msgbox("Goodbye!")
	If IsEmpty(ExitVal) Then
	'operation cancelled
    Msgbox ExitVal
Else
    'something has entered even zero-length
    MsgBox retVal
End If
