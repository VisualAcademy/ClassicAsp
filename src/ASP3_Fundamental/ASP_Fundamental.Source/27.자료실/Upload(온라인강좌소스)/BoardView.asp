<%
'����
Dim objCon: Dim objCmd: Dim objRs
'Ŀ�ؼ�
Set objCon = Server.CreateObject("ADODB.Connection")
'����
objCon.Open(Application("CONNECTION_STRING"))
	'Ŀ���
	Set objCmd = Server.CreateObject("ADODB.Command")
	'��Ƽ��Ŀ�ؼ�
	objCmd.ActiveConnection = objCon
	'Ŀ��� �ؽ�Ʈ 1
	objCmd.CommandText = "Update Upload Set ReadCount = ReadCount + 1 Where Num = " & Request("Num")
	'��� ����
	objCmd.Execute()
		'Ŀ��� �ؽ�Ʈ 2
		objCmd.CommandText = "Select * From Upload Where Num = " & Request("Num")
		'��� ���� �� ���ڵ�� ��ȯ
		Set objRs = objCmd.Execute()
			'��� ����� ���
			Call ShowRecordSet(objRs)
		'���ڵ�� �ݱ�
		objRs.Close()
		'���ڵ�� ����
		Set objRs = Nothing
	'����
	Set objCmd = Nothing
'�ݱ�
objCon.Close()
'����
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>
	<table border="1" width="100%">
<%
	If objRs.BOF Or objRs.EOF Then
%>
	<tr><td>�Էµ� �����Ͱ� �����ϴ�.</td></tr>
<%		
	Else
%>
	<tr><td>���� : </td><td><%=objRs("Title")%></td></tr>
	<tr><td>�ۼ��� : </td><td><%=objRs("Name")%></td></tr>
	<tr><td>�ۼ��� : </td><td><%=objRs("PostDate")%></td></tr>
	<tr><td colspan="2">
	<%
	Dim strContent: strContent = objRs("Content")
	'Text�� ����
	If objRs("Encoding") = "Text" Then
		'HTML �±� ����
		strContent = Replace(strContent, "&", "&amp;")
		strContent = Replace(strContent, "<", "&lt;")
		strContent = Replace(strContent, ">", "&gt;")
		'���� ó��
		strContent = Replace(strContent, Chr(13) & Chr(10), "<br>")
		'strContent = Replace(strContent, vbCrLf, "<br>")
	ElseIf objRs("Encoding") = "Mixed" Then
		'���� ó��
		strContent = Replace(strContent, Chr(13) & Chr(10), "<br>")
	Else
		'//strContent = objRs("Content")
	End If	
	'���
	Response.Write(strContent)
	%>	
	</td></tr>
<%
	End If
%>
	<tr><td colspan="2" align="center"><a href="BoardList.asp">����Ʈ</a></td></tr>
	</table>
<%
End Sub
%>














