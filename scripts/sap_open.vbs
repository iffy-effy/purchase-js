' SAP 로그인 화면 열기 (비밀번호는 사용자가 직접 입력)
' 출력: loggedin=N (이미 로그인됨) / opened (로그인 화면 띄움) / ERR ...
CONN_NAME = "SAP Logon 항목 이름"   ' SAP Logon에 보이는 접속 항목 이름으로 바꾸세요
LNK_NAME = "SAP_자동로그인.lnk"   ' (선택) 비밀번호를 넣은 명령줄 바로가기 파일 이름
On Error Resume Next
Set sh = CreateObject("WScript.Shell")
Set app = Nothing
Set gui = GetObject("SAPGUI")
If Err.Number <> 0 Then
  Err.Clear
  sh.Run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe""", 1, False
  For k = 1 To 30
    WScript.Sleep 1000
    Err.Clear
    Set gui = GetObject("SAPGUI")
    If Err.Number = 0 Then Exit For
  Next
End If
Err.Clear
Set app = gui.GetScriptingEngine
If Err.Number <> 0 Or app Is Nothing Then WScript.Echo "ERR SAP Logon 실행 실패" : WScript.Quit
If app.Children.Count > 0 Then
  Set c = app.Children(CLng(0))
  If c.Children.Count > 0 Then
    If c.Children(CLng(0)).Info.User <> "" Then WScript.Echo "loggedin=" & app.Children.Count : WScript.Quit
  End If
  WScript.Echo "opened (로그인 화면이 이미 떠 있음)" : WScript.Quit
End If
' 비밀번호가 든 명령줄 바로가기가 있으면 그것으로 자동 로그인(파일 내용은 읽지 않음)
lnk = sh.ExpandEnvironmentStrings("%USERPROFILE%") & "\SAP바로가기\" & LNK_NAME
If CreateObject("Scripting.FileSystemObject").FileExists(lnk) Then
  sh.Run """" & lnk & """", 1, False
  For k = 1 To 40
    WScript.Sleep 1000
    Err.Clear
    If app.Children.Count > 0 Then
      Set c = app.Children(CLng(0))
      If c.Children.Count > 0 Then
        If c.Children(CLng(0)).Info.User <> "" Then WScript.Echo "autologin" : WScript.Quit
      End If
    End If
  Next
  WScript.Echo "opened (자동 로그인 실패 - 로그인 화면 확인 필요)" : WScript.Quit
End If
Err.Clear
Set c = app.OpenConnection(CONN_NAME, True)
If Err.Number <> 0 Then WScript.Echo "ERR 접속 실패: " & Err.Description : WScript.Quit
WScript.Echo "opened"
