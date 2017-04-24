'Data: 2017. 02. 20
'Author: LeeJunHwan
'Content: AhnTree Auto Login Script for IE

WScript.Quit Main

Function Main
  Set IE = WScript.CreateObject("InternetExplorer.Application", "IE")
  IE.Visible = True
  IE.Navigate "http://10.2.102.150/AhnTree/"
  Wait IE

  With IE.Document
    .all.id.value = "jhlee.trainee" 'ID
    .all.pass.value = "123123" 'Password
    .all.form.submit()
  End With

  intTimeout = 1
  strMessage = "Welcome jhlee.trainee" 'Welcome Message
  strTitle = "Hello, Junan!"

  Set WshShell = CreateObject("WScript.Shell")
  intResult = WshShell.Popup(strMessage, intTimeout, strTitle)

  IE.Navigate "http://10.2.102.150/AhnTree/attendance.page"
End Function

Sub Wait(IE)
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy 
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy 
End Sub
