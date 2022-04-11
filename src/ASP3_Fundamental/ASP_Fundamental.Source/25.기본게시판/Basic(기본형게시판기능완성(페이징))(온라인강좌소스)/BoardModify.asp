<%
Dim objCon: Dim objRs: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
	'Set objRs = objCon.Execute("Select * From Basic Where Num = " & Request("Num"))
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	'레코드셋.Open SQL문, 커넥션객체, 커서타입, 락타입, 옵션
	strSql = "Select * From Basic Where Num = " & Request("Num")
	objRs.Open strSql, objCon, 3
	'출력
	Call ShowRecordSet(objRs)
	objRs.Close()
	Set objRs = Nothing
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
	If objRs.BOF Or objRs.EOF Then
	Else
%>
		<pre>
		<form action="./BoardModifyProcess.asp" method="post">
		<input type="hidden" name="Num" value="<%=Request("Num")%>">
		이름: <input type="text" name="Name" value="<%=objRs("Name")%>">
		이메일:<input type="text" name="Email" value="<%=objRs("Email")%>">
		홈페이지:<input type="text" name="Homepage" value="<%=objRs("Homepage")%>">
		제목:<input type="text" name="Title" value="<%=objRs("Title")%>">
		내용:<textarea name="Content" cols=40 rows=10><%=objRs("Content")%></textarea>
		인코딩:
			<select name="Encoding">
				<option value="Text">Text</option>
				<option value="HTML">HTML</option>
				<option value="Mixed">Mixed</option>
			</select>
		비밀번호<input type="password" name="Password">
		<input type="submit" value="수정">
		</form>
		</pre>
<%
	End If
End Sub
%>