<%
    ' 페이지 버퍼링 하려면, for IIS 4.0
   Response.Buffer = true
%>
<br> 1. HTML 스트림입니다.
<br> 2. HTML 스트림입니다. 우선 이 내용을 보냅니다.
<%
    ' 현재 버퍼에 만들어진 페이지를 클라이언트로 보냅니다.
    Response.Flush 
%>
<br> 3. 다음 HTML 스트림입니다. 
<br> 4. 다음 HTML 스트림입니다. 이 내용은 보내지 않습니다.
<%
    ' 현재 버퍼에 만들어진 내용을 삭제합니다.
    Response.Clear 
%>
<%  
    Response.Write "<br> 5. Response.Write 를 이용해도 마찬가지..."
    ' 현재 버퍼에 만들어진 페이지를 클라이언트로 보냅니다.
    Response.Flush
%>
<br> 6. HTML 스트림입니다.
<br> 7. 이것이 마지막 HTML 스트림입니다.
<%
    ' 현재 버퍼에 만들어진 페이지를 클라이언트로 보내고 페이지를 종료합니다.
    Response.End 
%>
<p> 8. 이 HTML 스트림은 클라이언트에 보내지 않습니다.