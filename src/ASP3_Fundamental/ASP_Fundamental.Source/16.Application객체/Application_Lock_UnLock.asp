<%
	Application.Lock   '변수 수정을 막는다.

	'Check라는 애플리케이션 변수를 선언하고, 이 문서를 읽을때마다
	'1증가시킨다.
	Application("Check") = Application("Check") + 1

	Application.UnLock   '변수의 락을 풀어준다.

	Session("Check") = Session("Check") + 1
%>
새로읽을 때마다 값이 증가(Application변수) : <%=Application("Check")%><br>
새로읽을 때마다 값이 증가(Session변수) : <%=Session("Check")%>






