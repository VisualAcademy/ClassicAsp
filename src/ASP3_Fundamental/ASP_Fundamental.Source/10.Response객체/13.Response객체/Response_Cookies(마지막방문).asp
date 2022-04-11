<%
    ' 페이지에 대한 캐시를 만들지 않는다.      
    'Response.Expires = 0

    ' 쿠키 값 읽어오기
    ' 이전에 방문한 시간을 Cookie 콜렉션에서 읽어온다.
    LastVisit = Request.Cookies("LastVisit")
 
    ' 쿠키 값 저장하기, 현재 날짜/시간 저장
    ' 이 정보는 다음에 방문했을 때 보여줄 마지막 방문 시간/날짜
    Response.Cookies("LastVisit") = FormatDateTime( Now )
%>

<h3> Cookie 이용 </h3>
<hr>
<%           
    If LastVisit = "" Then
        ' 처음 읽은 경우 쿠키 정보가 없음 
        Response.Write "처음 온 것을 환영합니다."
    Else
        ' 이미 읽은 적이 있는 경우 마지막 방문 날짜/시간을 보여줌
        Response.Write "당신은 " & LastVisit  & "에 방문한 적이 있습니다."
    End If 
%>
<p><A href="Response_Cookies(마지막방문).asp"> 페이지 다시 읽어오기 </A>

