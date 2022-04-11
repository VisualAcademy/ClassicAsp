
관리자 메모 지우기 폼<br>

<%=Request("Num")%>번 메모를 삭제하시겠습니까?<br>

<form action="./MemoDeleteProcess.asp" method="post">

<input type="hidden" name="Num" value="<%=Request("Num")%>">

관리자 패스워드 : <input type="password" name="Password"><br>

<input type="submit" value="전송">

</form>
