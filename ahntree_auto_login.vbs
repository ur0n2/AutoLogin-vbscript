'Data: 2017. 02. 20
'Author: LeeJunHwan
'Content: AhnTree Auto Login Script for IE

WScript.Quit Main

Function Main
  Set IE = WScript.CreateObject("InternetExplorer.Application", "IE")
  IE.Visible = True
  IE.Navigate "http://x.x.x.x" 'Base URL
  Wait IE

  With IE.Document
    .all.id.value = "TEST_ID" 'ID
    .all.pass.value = "123123" 'Password
    .all.form.submit()
  End With

  ' Messagebox Create
  intTimeout = 1
  strMessage = "Welcome jhlee.trainee" 'Welcome Message
  strTitle = "Hello, Junan!"

  Set WshShell = CreateObject("WScript.Shell")
  intResult = WshShell.Popup(strMessage, intTimeout, strTitle)

  IE.Navigate "http://x.x.x.x/attendance.page" ' View page after auto-login
End Function

Sub Wait(IE)
  Do
    WScript.Sleep 500 '0.5second
  Loop While IE.ReadyState < 4 And IE.Busy 
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy 
End Sub
