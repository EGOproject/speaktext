Dim Message, Speak
Message= Inputbox("What would you want EGO to say??", "EGO AI")
Set Speak= Createobject ("Sapi.Spvoice")
Speak.Speak Message
