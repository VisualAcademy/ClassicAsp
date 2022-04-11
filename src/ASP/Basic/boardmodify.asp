<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boardmodify.asp
' Program Description : 수정하기 폼 == 글쓰기 폼과 비슷
' QueryString : 반드시 view.asp에서 Num=1 식으로 전송
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
Option Explicit
Dim Num: Num = Request("Num")
Dim objCon: Dim objRs: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))
Set objRs = Server.CreateObject("ADODB.RecordSet")
strSql = "Select * From Basic Where Num = " & Num
objRs.Open strSql, objCon
If objRs.BOF Or objRs.EOF Then
	Response.Write("수정할 페이지가 잘못되었습니다.")
Else
	Call ShowModifyForm(objRs)
End If
objRs.Close(): objCon.Close()
Set objRs = Nothing: Set objCon = Nothing
%>
<%Sub ShowModifyForm(objRs)%>
<form action="./boardmodifyprocess.asp" method="post">
<pre>
<input type="hidden" name="Num" value="<%=Num%>">
이름 : <input type="text" name="Name" value="<%=objRs("Name")%>">
이메일 : <input type="text" name="Email" value="<%=objRs("Email")%>">
홈페이지 : <input type="text" name="Homepage" value="<%=objRs("Homepage")%>">
제목 : <input type="text" name="Title" value="<%=objRs("Title")%>">
내용 : <textarea name="Content" cols="40" rows="10"><%=objRs("Content")%></textarea>
인코딩 : <select name="Encoding"><option value="Text">Text</option><option value="HTML">HTML</option></select>
패스워드 : <input type="password" name="Password">
<input type="submit" value="작성 완료">
</pre>
</form>
<%End Sub%>