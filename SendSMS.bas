' from http://dwarf1711.blogspot.com/2007/10/vbscript-urlencode-function.html 
Function URLEncode(ByVal str)
 Dim strTemp, strChar
 Dim intPos, intASCII
 strTemp = ""
 strChar = ""
 For intPos = 1 To Len(str)
  intASCII = Asc(Mid(str, intPos, 1))
  If intASCII = 32 Then
   strTemp = strTemp & "+"
  ElseIf ((intASCII < 123) And (intASCII > 96)) Then
   strTemp = strTemp & Chr(intASCII)
  ElseIf ((intASCII < 91) And (intASCII > 64)) Then
   strTemp = strTemp & Chr(intASCII)
  ElseIf ((intASCII < 58) And (intASCII > 47)) Then
   strTemp = strTemp & Chr(intASCII)
  Else
   strChar = Trim(Hex(intASCII))
   If intASCII < 16 Then
    strTemp = strTemp & "%0" & strChar
   Else
    strTemp = strTemp & "%" & strChar
   End If
  End If
 Next
 URLEncode = strTemp
End Function

' A function to send an SMS message using the eztexting.com API 
' Requires a file auth.txt in the same directory (?) containing the username and password for the EZTexting API on separate lines
Public Function SendSMS(ByVal MessageDest, ByVal MessageSubject, ByVal MessageBody)
    Dim xhr
    Dim AccountUsername
    Dim AccountPassword
    Dim objFSO, objFile
    Dim postParams
    Dim res, code
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If objFSO.FileExists("auth.txt") = False Then
        MsgBox("Could not find auth.txt!")
        Exit Function
    End If

    Set objFile = objFSO.OpenTextFile("auth.txt", 1)

    AccountUsername = URLEncode(objFile.ReadLine)

    If objFile.AtEndOfStream = True Then
        MsgBox("Could not read password from auth.txt!")
        Exit Function
    End If
    AccountPassword = URLEncode(objFile.ReadLine)

    Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    postParams = "User=" & URLEncode(AccountUsername) & _
                 "&Pass=" & URLEncode(AccountPassword) & _
                 "&PhoneNumber=" & MessageDest & _
                 "&Subject=" & URLEncode(MessageSubject) & _
                 "&Message=" & URLEncode(MessageBody)

    URL = "https://app.eztexting.com/api/sending/"
    
    xhr.Open "POST", URL, False
    xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    xhr.send postParams

    ' xhr.responseText now holds the response as a string.
    res = xhr.responseText
    code = xhr.Status

    If code <> 200 Then
        MsgBox("Server replied with status code " & code)
        SendSMS = code
        Exit Function
    End If

    If res = "-1" Then
        MsgBox("Authentication problem, or API access not allowed")
    ElseIf res = "-2" Then
        MsgBox("Credit limit reached")
    ElseIf res = "-5" Then
        MsgBox("Recipient is on the local opt-out list")
    ElseIf res = "-7" Then
        MsgBox("Invalid message or subject (too long or contains invalid characters)")
    ElseIf res = "-104" Then
        MsgBox("Recipient is on the global opt-out (does not receive messages from ANY EZTexting account")
    ElseIf res = "-106" Then
        MsgBox("Incorrectly formatted phone number, must be 10 digits")
    ElseIf res = "-10" Then
        MsgBox("Server replied that it experienced an unknown error")
    ElseIf res <> "1" Then
        MsgBox("Unknown response from server")
    End If

    SendSMS = xhr.responseText
End Function
