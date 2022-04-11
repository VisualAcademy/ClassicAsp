<%
Dim objCon: Dim strSql: Dim objRs
Dim Num: Num = Request("Num")

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))
	strSql = "Select * From Upload Where Num = " & Num
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	objRs.Open strSql, objCon
		'각각의 폼에 이미 입력된 값을 모양만들어 출력
		Call ShowRecordSet(objRs)
	objRs.Close()
	Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
	If objRs.BOF Or objRs.EOF Then'//뒤로 돌리거나 리스트로 이동
	Else'BoardWrite.asp에서 작성한 폼의 value에 필드값 세팅
%>
<pre><form action="./BoardModifyProcess.asp" method="post" ID="Form1">
<input type="hidden" name="Num" value="<%=objRs("Num")%>">
이름: <input type="text" name="Name" value="<%=objRs("Name")%>" ID="Text1">
이메일:<input type="text" name="Email" value="<%=objRs("Email")%>" ID="Text2">
홈페이지:<input type="text" name="Homepage" ID="Text3" value="<%=objRs("Homepage")%>">
제목:<input type="text" name="Title" ID="Text4" value="<%=objRs("Title")%>">
내용:<textarea name="Content" cols=40 rows=10 ID="Textarea1"><%=objRs("Content")%></textarea>
인코딩:<select name="Encoding" ID="Select1">
		<option value="Text" 
			<%If objRs("Encoding") = "Text" Then%>
			selected
			<%End If%>
		>Text</option>
		<option value="HTML"
		<%If objRs("Encoding") = "HTML" Then%>selected<%End If%>
		>HTML</option>
		<option value="Mixed"
		<%If objRs("Encoding") = "Mixed" Then%>selected<%End If%>
		>Mixed</option>
	</select>
<!--자료실에서 추가-->
비밀번호<input type="password" name="Password" ID="Password1">
<input type="submit" value="전송" ID="Submit1" NAME="Submit1">
</form></pre>
<%
	End If
End Sub
%>








