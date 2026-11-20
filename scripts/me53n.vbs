' 사용법: cscript //nologo me53n.vbs <PR번호> check|workflow
'  check    : ME53N에서 PR을 열고 제목의 PR번호를 출력
'  workflow : PR을 열어 번호가 맞을 때만 Services for Object ▼ > Start Workflow 실행
'             (SAP 보안 창에서 멈추므로 반드시 백그라운드로 실행)
On Error Resume Next
pr = WScript.Arguments(0) : mode = LCase(WScript.Arguments(1))
Set app = GetObject("SAPGUI").GetScriptingEngine
If app.Children.Count = 0 Then WScript.Echo "ERR no SAP connection" : WScript.Quit 1
Set s = app.Children(CLng(0)).Children(CLng(0))

ok = False
For tries = 1 To 2
  s.findById("wnd[0]/tbar[0]/okcd").Text = "/nME53N"
  s.findById("wnd[0]").sendVKey 0
  If s.findById("wnd[1]", False) Is Nothing Then s.findById("wnd[0]").sendVKey 17
  s.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_B").Select
  s.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").Text = pr
  s.findById("wnd[1]/tbar[0]/btn[0]").press
  If Not s.findById("wnd[1]", False) Is Nothing Then s.findById("wnd[1]").sendVKey 12
  t = s.findById("wnd[0]").Text
  If InStr(t, pr) > 0 Then ok = True : Exit For
Next
WScript.Echo "title=" & s.findById("wnd[0]").Text
If Not ok Then WScript.Echo "ERR PR number mismatch (expected " & pr & ") sbar=" & s.findById("wnd[0]/sbar").Text : WScript.Quit 2
If mode <> "workflow" Then WScript.Quit 0

Set g = s.findById("wnd[0]/titl/shellcont/shell")
g.pressContextButton "%GOS_TOOLBOX"
WScript.Echo "starting workflow (click Allow on the SAP GUI Security prompt, without Remember My Decision)"
g.selectContextMenuItemByText "Start Workflow"
WScript.Echo "workflow returned err=" & Err.Number & " " & Err.Description
