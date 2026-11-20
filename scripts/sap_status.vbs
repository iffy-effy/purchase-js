' SAP GUI 로그인/세션 상태 확인
On Error Resume Next
Set app = GetObject("SAPGUI").GetScriptingEngine
If Err.Number <> 0 Then WScript.Echo "connections=0 (SAP GUI 없음)" : WScript.Quit
WScript.Echo "connections=" & app.Children.Count
For i = 0 To app.Children.Count - 1
  Set c = app.Children(CLng(i))
  For j = 0 To c.Children.Count - 1
    Set s = c.Children(CLng(j))
    WScript.Echo s.Id & " tx=" & s.Info.Transaction & " popup=" & (s.Children.Count > 1)
  Next
Next
