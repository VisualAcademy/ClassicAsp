<%
Dim objCon: Dim objRs: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
	'Set objRs = objCon.Execute("Select * From Basic Where Num = " & Request("Num"))
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	'���ڵ��.Open SQL��, Ŀ�ؼǰ�ü, Ŀ��Ÿ��, ��Ÿ��, �ɼ�
	strSql = "Select * From Basic Where Num = " & Request("Num")
	objRs.Open strSql, objCon, 3
	'���
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
		�̸�: <input type="text" name="Name" value="<%=objRs("Name")%>">
		�̸���:<input type="text" name="Email" value="<%=objRs("Email")%>">
		Ȩ������:<input type="text" name="Homepage" value="<%=objRs("Homepage")%>">
		����:<input type="text" name="Title" value="<%=objRs("Title")%>">
		����:<textarea name="Content" cols=40 rows=10><%=objRs("Content")%></textarea>
		���ڵ�:
			<select name="Encoding">
				<option value="Text">Text</option>
				<option value="HTML">HTML</option>
				<option value="Mixed">Mixed</option>
			</select>
		��й�ȣ<input type="password" name="Password">
		<input type="submit" value="����">
		</form>
		</pre>
<%
	End If
End Sub
%>