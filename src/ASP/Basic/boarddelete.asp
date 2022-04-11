<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boarddelete.asp
' Program Description : 지우기 폼
' QueryString : 반드시  view.asp에서 Num=1 식으로 전송
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
Dim Num: Num = Request("Num")
%>
<%=Num%>번 글을 삭제하시겠습니까?<br>
<form action="./boarddeleteprocess.asp"
method="post">
<!--번호 : --><input type="hidden" name="Num" 
		value="<%=Num%>"><br>
현재글 비밀번호 : <input type="password" name="Password">
<br>
<input type="submit" value="삭제하기">
<input type="button" value="취소" onclick="history.go(-1);">
</form>