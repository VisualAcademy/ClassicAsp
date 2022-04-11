<%
    ' 페이지 버퍼링 설정(기본값)
   Response.Buffer = True
%>
	1. 이 내용은 출력됩니다.<br>
<%
    ' 버퍼에 있는 내용 출력
    Response.Flush 
%>
	2. 이 내용은 아래 Clear메서드에 의해 출력되지 않습니다.<br>
<%
    ' 버퍼에 있는 내용 삭제
    Response.Clear 
%>
<%  
    Response.Write "3. 이 내용도 출력되지 않음<br>"
    ' Response.Write()메서드도 Clear()메서드에 영향을 받음.
    Response.Clear
%>
	4. 아래 End메서드에 의해서 여기까지만 실행됩니다.<br> 
<%
    ' 버퍼에 있는 내용 출력 후 페이지 종료
    Response.End 
%>
5. 이 내용은 출력되지 않습니다.<br> 