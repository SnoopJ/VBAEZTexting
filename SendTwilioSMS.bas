Public Function SendTwilioMessage(ByVal MessageDest, ByVal MessageBody)
    Dim xhr
    Dim AccountSid
    Dim AuthToken
    Dim MessageSrc
    Dim postParams
    
    AccountSid = ""  ' Your account SID here
    AuthToken = ""   ' Your auth token here.  This is a secret, DO NOT DISTRIBUTE!
    MessageSrc = ""  ' Your "from" number here.
    
    Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' Twilio API endpoint for sending an SMS:
    ' POST to https://api.twilio.com/2010-04-01/Accounts/AccountSid/Messages
    
    postParams = "From=" & MessageSrc & "&To=" & MessageDest & "&Body=" & MessageBody
    URL = "https://api.twilio.com/2010-04-01/Accounts/" & AccountSid & "/Messages.json"
    
    xhr.Open "POST", URL, False
    xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xhr.SetCredentials AccountSid, AuthToken, 0
    
    xhr.send postParams
    ' xhr.responseText now holds the response as a string.
    
    MsgBox (xhr.responseText)
End Function
