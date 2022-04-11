<%
    Response.Write("현재 스크립트 타임아웃 시간 : " & _
		Server.ScriptTimeout & "초<p>")

    ' 스크립트 타임아웃 시간을 3분으로 바꿈
    Server.ScriptTimeout = 180

    Response.Write("바뀐 스크립트 타임아웃 시간 : " & _
		Server.ScriptTimeout & "초<p>")
%>