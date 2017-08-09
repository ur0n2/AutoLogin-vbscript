'Date: 2017. 08. 01
'Author: LeeJunHwan
'Content: ACE Auto Login Script for IE

WScript.Quit Main

Function Main
  Set IE = WScript.CreateObject("InternetExplorer.Application", "IE")
  IE.Visible = True
  IE.Navigate "https://ace.ahnXXX.com"
  Wait IE
  With IE.Document
    .all.txtid.value = "XXXXX" 'ID
    .all.txtPWD.value = "123123" 'Password
    .all("btnLogin").Click
  End With


End Function

Sub Wait(IE)
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy 
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy 
End Sub
