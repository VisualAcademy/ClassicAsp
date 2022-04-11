<%
'[1] Private한 전역변수
Session("Now") = Now()
%>

요청시간:<%= Session("Now")%><br />
고유키값:<%= Session.SessionID %><br />
