<%
Dim objCon: Dim strSql: Dim objRs
Dim Num: Num = Request("Num")

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))
	strSql = "Select * From Upload Where Num = " & Num
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	objRs.Open strSql, objCon
		'������ ���� �̹� �Էµ� ���� ��縸��� ���
		Call ShowRecordSet(objRs)
	objRs.Close()
	Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
	If objRs.BOF Or objRs.EOF Then'//�ڷ� �����ų� ����Ʈ�� �̵�
	Else'BoardWrite.asp���� �ۼ��� ���� value�� �ʵ尪 ����
%>
<pre><form action="./BoardModifyProcess.asp" method="post" ID="Form1">
<input type="hidden" name="Num" value="<%=objRs("Num")%>">
�̸�: <input type="text" name="Name" value="<%=objRs("Name")%>" ID="Text1">
�̸���:<input type="text" name="Email" value="<%=objRs("Email")%>" ID="Text2">
Ȩ������:<input type="text" name="Homepage" ID="Text3" value="<%=objRs("Homepage")%>">
����:<input type="text" name="Title" ID="Text4" value="<%=objRs("Title")%>">
����:<textarea name="Content" cols=40 rows=10 ID="Textarea1"><%=objRs("Content")%></textarea>
���ڵ�:<select name="Encoding" ID="Select1">
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
<!--�ڷ�ǿ��� �߰�-->
��й�ȣ<input type="password" name="Password" ID="Password1">
<input type="submit" value="����" ID="Submit1" NAME="Submit1">
</form></pre>
<%
	End If
End Sub
%>








