---
name: purchase-js
description: {우리 팀} SAP 구매요청(PR)을 품목 문장 하나로 진행한다. 엑셀 PR자동화 파일에 품목을 넣고 ME51N 저장 → ME53N Start Workflow → 전자결재(joyNABI) 구매요청 작성 창을 크롬에서 열어 제목과 개인결재선(구매팀/기술팀)까지 넣고 기안 직전에 멈춘다(기안은 사용자가 직접). "소포제 100kg, 고무이중코팅장갑 20개 구매팀으로 구매요청 해줘", "무전기 수리 2개 기술팀으로 PR 올려줘", "견적서 첨부해서 구매요청", "ME51N 올려줘"처럼 품목과 수량을 말하며 구매요청·PR·발주요청·구매 품의를 부탁하면, 스킬 이름을 말하지 않아도 반드시 이 스킬을 쓴다.
---

# SAP 구매요청 → 전자결재(joyNABI) 작성 창(제목·결재선까지)

사용자({우리 팀}, {팀코드})가 품목 문장을 주면 아래 순서로 진행한다. 이 일은 실제 회사 시스템에 PR을 만든다. 그래서 **품목이 확실할 때만 저장**한다. 전자결재(joyNABI)에서는 **제목과 결재선(개인결재선)까지 넣고 멈춘다**. [기안]은 사용자가 직접 누른다.

## 고정 정보

| 항목 | 값 |
|---|---|
| 엑셀 파일 | `{엑셀 파일 경로}\ME51N_현장소모품_PR자동화.xlsm` |
| 매크로 모듈 | 같은 폴더 `ME51N_PR.bas` (PR_ClearInput / PR_AddRow / PR_Run / Create_PR_ME51N) |
| 스크립트 | 이 스킬의 `scripts/` (PowerShell은 UTF-8 BOM으로 저장되어 있음) |
| 전자결재(joyNABI) 작성 창 주소 | `https://{결재시스템 주소}/approval/legacy/goFormLink.do?IV_WFKEY=PR01{PR번호}001300&logonId={로그온ID}&fmpf=WF_FORM_REQUEST_PR&mode=DRAFT` |

## 1. 요청 문장 해석

- 품목마다 `품목 / 수량 / 단위(있으면)`로 나눈다. 예: "소포제 100kg, 고무이중코팅장갑 20개, 무전기 수리 2개".
- **구매팀/기술팀**: 문장에 나온 쪽을 쓴다. 없으면 시작 전에 물어본다(제목 끝 괄호가 달라지므로 추측하지 않는다).
- **납기**: 말하지 않으면 오늘 날짜(yyyy-mm-dd).
- **견적서**: 파일 경로를 주면 3단계에서 첨부한다.
- **비고**: 사용자가 따로 말한 사용처가 있으면 그 문구(35자 이내). 없으면 비워 둔다(엑셀이 품목마스터의 최근 요청 사유를 자동으로 넣는다).

## 2. 품목 확정 → 엑셀 → ME51N 저장

1. **엑셀 열기**: 파일이 안 열려 있으면 `Start-Process`로 연다. 열린 직후엔 COM에 안 잡히니(MK_E_UNAVAILABLE) 5초 간격으로 최대 1분 `GetActiveObject`를 다시 시도한 뒤 진행한다. 닫혀 있는 상태에서 COM으로 새로 열지 말 것 — 매크로 보안 때문에 사용자가 연 창을 쓰는 게 안전하다.
2. **SAP 로그인 확인**: `cscript //nologo scripts\sap_open.vbs`
   - `loggedin=N` → 이미 로그인됨, 다음 단계로.
   - `autologin` → `%USERPROFILE%\SAP바로가기\SAP_자동로그인.lnk`(사용자가 비밀번호를 넣어 둔 명령줄 바로가기)로 자동 로그인 완료. 다음 단계로. 이 파일은 실행만 하고 절대 열어 보거나 내용을 읽지 않는다.
   - `opened (자동 로그인 실패...)` → 바로가기는 실행됐지만 로그인이 안 됨(비밀번호 변경 등). 사용자에게 로그인과 바로가기 비밀번호 갱신을 부탁한다.
   - `opened` → SAP Logon을 켜고 SAP 시스템 로그인 화면을 띄운 것. 사용자에게 "SAP 로그인 화면을 열었습니다. ID·비밀번호를 입력해 주세요"라고 알리고 기다린다. 사용자가 됐다고 하면 `sap_status.vbs`로 `tx=`가 S000이 아닌지(로그인 완료) 확인한다.
   - `ERR ...` → SAP Logon을 직접 실행해 달라고 부탁한다.
   - ID·비밀번호는 절대 대신 입력하지 않는다. 로그인 창은 여러 번 띄우지 않는다(스크립트가 이미 떠 있으면 그대로 둔다).
3. **품목 찾기**: 품목마다
   `powershell -File scripts\find_item.ps1 -Query "고무이중코팅장갑"`
   - 결과 1건 → 그대로 확정.
   - 여러 건 → 상위 후보(요청 횟수 순, 번호·품목명·자재코드·단위)를 채팅으로 보여 주고 고르게 한다. 단어를 더 주면 다시 검색한다.
   - 0건 → 다른 표현을 물어본다. 품목마스터에 없는 품목이면 마스터에 추가가 필요하다고 알리고 그 품목은 빼고 진행할지 묻는다.
   - 수량 단위가 마스터 단위와 다르면(예: "100kg"인데 마스터는 EA/PAIL) 어떻게 환산할지 확인한다. 임의로 바꾸지 않는다.
4. **엑셀 입력**: `scripts\excel_pr.ps1 -Action clear` 후 품목마다
   `scripts\excel_pr.ps1 -Action add -Item "<마스터 품목명 정확히>" -Qty 10 -Due 2026-11-20 [-Note "…"]`
   - 결과가 `ERR`이면 멈추고 이유를 알린다.
   - 원재료는 지정업체(기본 Y)와 "({팀코드}) yyyy년 mm월 원재료 구매요청" 비고가 자동으로 들어간다.
5. **입력 내용 보고**: 저장 전에 품목·수량·납기·자재코드·코스트센터·G/L을 표로 한 번 보여 준다. 사용자가 이미 품목을 정확히 지정했으면 확인을 기다리지 않고 진행해도 된다(이 스킬을 부른 것 자체가 저장 요청이다).
6. **ME51N 저장**: `scripts\excel_pr.ps1 -Action run -Save`
   - 반환 문자열에서 `Purchase requisition number 10xxxxxxxx created` 의 번호가 **PR번호**다. P열과 '구매요청이력' 시트에도 자동 기록된다.
   - 번호가 없으면(오류 메시지) 멈추고 그대로 전한다.

## 3. ME53N → (견적서) → Start Workflow

1. `cscript //nologo scripts\me53n.vbs <PR번호> check` → `title=Display Purchase Req. <PR번호>` 인지 확인.
   ME53N은 마지막에 본 PR을 먼저 띄우므로, **제목의 번호가 ME51N에서 받은 번호와 다르면 절대 다음 단계로 가지 않는다**(다시 실행, 그래도 다르면 멈춤). 번호가 엇갈리면 전자결재(joyNABI)에서 오류가 난다.
2. **견적서가 있으면**: ME53N 왼쪽 위 'Services for Object' 아이콘 옆 ▼ → **Attatchment Management**(SAP 표기 그대로)에서 파일을 추가한다. 이 화면은 아직 스크립트로 검증하지 않았으니, 처음 한 번은 화면을 보며(computer-use) 진행하고 결과를 사용자에게 확인받는다.
3. **Start Workflow**: `scripts\me53n.vbs <PR번호> workflow` 를 **백그라운드로** 실행한다(SAP이 보안 창에서 멈춰 스크립트가 끝나지 않는다).
4. 'SAP GUI Security' 창이 뜨면 computer-use로 **[Allow]** 를 누른다. **"Remember My Decision"은 체크하지 않는다**(체크하면 전자결재(joyNABI) 연결이 계속 실패한다 — 사용자 확인 사항). 사용자도 Allow를 눌러도 된다고 허락했다.
5. 기본 브라우저(크롬)에도 전자결재(joyNABI) 작성 창이 하나 뜬다. 그 탭은 에이전트 탭 그룹 밖이라 쓰지 않는다. 사용자에게 "SAP이 띄운 창은 닫아도 된다"고 알린다(두 창에서 각각 기안하면 중복 상신되므로 꼭 안내).

## 4. 전자결재(joyNABI) 작성 창 (크롬 — Claude in Chrome)

크롬 도구(`mcp__claude-in-chrome__*`)를 ToolSearch로 한 번에 불러온다. `list_connected_browsers`가 비어 있으면 사용자에게 크롬 확장 연결을 부탁한다.

1. `tabs_context_mcp {createIfEmpty:true}` → 그 탭에서 위 주소(PR번호 넣은 goFormLink)로 `navigate`. `approval_Form.do...` 로 넘어가면 성공(logonId로 자동 로그인).
   - 화면이 비고 `"FAIL"` 이면 문서 키 뒷자리(`001300`)가 다른 것이다. SAP이 크롬에 띄운 창의 주소를 사용자에게 복사해 달라고 부탁하고, 받은 주소로 연다. 새 뒷자리는 이 문서에 반영한다.
2. `get_page_text`로 구매내역(PR번호·품목·수량)이 이번 PR과 같은지 확인한다.
3. **제목**(`#Subject`)을 지우고 새로 입력한다(javascript_tool로 value 설정 후 input/change 이벤트):
   - 품목 1~2개: `[ {PR번호} ] {품목} 구매요청 결재건 ({구매팀|기술팀})` (2개면 쉼표로)
   - 품목 3개 이상: `[ {PR번호} ] 현장소모품({품목1}, {품목2}, …) 구매요청 결재건 ({팀})`
   - 품목명은 괄호 속 규격을 뺀 짧은 이름. 입력 후 값을 다시 읽어 확인한다.
4. **결재선** — 팝업 창은 탭 그룹 밖이라 직접 못 잡는다. 같은 출처라 **작성 창 탭의 javascript_tool에서 팝업 window를 이름으로 잡아** 조작한다. 각 단계는 따로 실행하고 결과를 확인한다.
   - (a) **팝업은 진짜 클릭으로 연다.** `find`로 상단 '결재선' 버튼 ref를 찾아 `computer left_click {ref}`. JS의 `.click()`이나 좌표 추정 클릭은 크롬 팝업 차단에 걸려 화면에 "팝업 차단을 허용해 주세요" 알림(#AlertLayer)이 뜬다. 뜨면 그 [확인] ref를 클릭해 닫고 다시 진짜 클릭한다.
   ```js
   // 팝업 잡기 준비 (작성 창 탭에서 한 번)
   if(!window.__origOpen) window.__origOpen = window.open;
   // (b) 3초 뒤 팝업 잡기 → 개인결재선 목록 열기
   const w = window.__origOpen.call(window,'','DivPop_'); const d = w.document;
   d.querySelector('#btnOpen').click(); await new Promise(r=>setTimeout(r,2500));
   // (c) 팀에 맞는 개인결재선 라디오 선택 → [적용]
   const name = '[발신]구매팀 구매';   // 기술팀이면 '[발신]기술팀 구매'
   const it = [...d.querySelectorAll('.preset_item')].find(x=>x.querySelector('.preset_title').innerText.trim()===name);
   const r = it.querySelector('input[type=radio]'); r.click();
   if(!r.checked){ r.checked=true; r.dispatchEvent(new Event('change',{bubbles:true})); }
   const dlg = [...d.querySelectorAll('[id^=dialog_]')].pop();
   [...dlg.querySelectorAll('.dialog_footer button')].find(b=>b.innerText.trim()==='적용').click();
   ```
   - (c) 뒤 `d.body.innerText`의 '결재자'~'결재선 미리보기' 부분을 읽어 **순서와 참조를 반드시 대조**한다. 다르면 [확인]을 누르지 말고 멈춰서 사용자에게 알린다.
     - 결재: {기안자}(기안) → {결재자1} → {결재자2} → {결재자3}
     - 참조(구매팀): {참조1}, {참조2}, {참조3}, {참조-구매팀4}
     - 참조(기술팀): {참조1}, {참조2}, {참조3}, {참조-기술팀4}, {참조-기술팀5}, {참조-기술팀6}
   - 맞으면 `w.document.querySelector('#btOK').click()` ([확인] — 작성 창에 결재선만 반영, 상신 아님). 팝업이 닫히고 작성 창 '부서결재' 표와 '참조'가 바뀐 것을 `document.body.innerText`로 확인한다.
   - 출력에 주소·쿼리 문자열이 섞이면 도구가 `[BLOCKED]`로 가린다. 필요한 텍스트만 반환하게 한다.
5. **여기서 멈춘다.** [기안]은 절대 누르지 않는다. 탭은 사용자가 기안할 수 있게 열어 둔다.

## 5. 마무리 보고

짧게 표로:
- PR번호, 품목/수량/금액, 바꾼 제목
- 적용한 결재선(결재 순서·참조) 요약
- "크롬의 Claude 탭 그룹에 있는 작성 창에서 내용 확인 후 [기안]만 누르시면 됩니다"
- 엑셀 이력 기록 여부, SAP이 띄운 중복 작성 창은 닫아 달라는 안내

## 문제가 생기면

| 증상 | 대응 |
|---|---|
| `SAP GUI 세션을 찾을 수 없습니다` / 중간에 로그아웃됨 | `sap_open.vbs`로 로그인 화면을 띄우고 사용자에게 로그인 요청 |
| 스크립트가 끝나지 않음 | SAP 화면에 팝업이 떠 있는지 스크린샷으로 확인. **SAP 스크립트 프로세스를 강제 종료하면 SAP 세션이 끊길 수 있으니** 팝업부터 처리한다 |
| ME51N 오류(계정지정 등) | PR_Run 반환 메시지와 엑셀 P열 내용을 그대로 전달. 품목마스터 K·L열 확인 안내 |
| 전자결재(joyNABI) 주소가 FAIL | 4-1 참고 (SAP이 띄운 크롬 창 주소 받기) |
| 크롬 도구 연결 안 됨 | 크롬 확장 설치·로그인 요청. 안 되면 제목까지만 하고 결재선은 사용자에게 안내 |
| 개인결재선 이름이 안 보임 | 팝업 목록 텍스트를 사용자에게 보여 주고 어떤 걸 쓸지 묻는다 |
| 엑셀이 응답 없음 | 셀 편집 중이거나 메시지창이 떠 있을 수 있다. 사용자에게 Esc를 부탁 |

## 하지 않을 것

- 전자결재(joyNABI) [기안] 클릭, 결재 상신, PR 삭제·변경. 개인결재선 삭제·'기본으로 설정'.
- SAP/전자결재(joyNABI) 비밀번호 입력, Edge 방문 기록 파일 직접 읽기.
- "Remember My Decision" 체크.
- 후보가 여러 개인데 임의로 고르기.
