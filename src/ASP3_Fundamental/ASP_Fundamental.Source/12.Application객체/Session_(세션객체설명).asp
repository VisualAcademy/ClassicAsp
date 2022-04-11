<!--
Session 객체
- 기본개념은 전체 웹사이트내에서 사용되는 전역변수(개인용(private))적 의미.
- 전역변수(개인용) : Session("변수이름")
   ex) Session("ID") = "관리자"
- .SessionID -> 고유한 세션번호 부여(각 사용자가 웹사이트에 들어올 때 OS가 자동부여)
- .Timeout -> 세션유지시간을 변경
   ex) Session.Timeout = 20   '세션유지시간을 20분으로 설정(기본값)
- .Abandon 메소드 : 걍 모든 세션을 쥑인다.(해당 웹사이트의 모든 세션값을 날린다...)

- Sesstion_OnStart 이벤트
- Sesstion_OnEnd 이벤트
-->
<%



%>

