<%
	'변수 선언
	Dim pwd

	'전송된 패스워드 값을 Request객체의 Form컬렉션으로 받아서 pwd변수에 저장
	pwd = Request.Form("pwd")

	'만약에 전송된 패스워드값이 1234(본인지정)이면 관리자 메시지 출력하고,
	'그렇지 않으면, 자바스크립트를 사용해서 이전 문서로 이동한다.
	If pwd = "1234" Then
		Response.Write("<center>관리자입니다.</center>")
	Else
%>
	<script language="JavaScript">
		window.alert("관리자가 아니군요\n다시 로그인하세요\n설마 해킹을???");
		history.go(-1);   //history.back();
	</script>
<%
	End If
%>
		