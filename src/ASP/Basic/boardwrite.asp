<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boardwrite.asp
' Program Description : 글쓰기 폼 페이지
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<form action="./boardwriteprocess.asp" method="post">
<pre>
이름 : <input type="text" name="Name">
이메일 : <input type="text" name="Email">
홈페이지 : <input type="text" name="Homepage">
제목 : <input type="text" name="Title">
내용 : <textarea name="Content" cols="40" rows="10"></textarea>
인코딩 : <select name="Encoding"><option value="Text">Text</option><option value="HTML">HTML</option></select>
패스워드 : <input type="password" name="Password">
<input type="submit" value="작성 완료">
</pre>
</form>