<%
'���� ���� ��
Dim objCon: Dim objRs: Dim strSql
'����¡ 1
Dim intPage: intPage = Request("Page")
If intPage = "" Then
	intPage = 1
End If
Dim intPageCount: intPageCount = 0
'Ŀ�ؼ� ��ü ����
Set objCon = Server.CreateObject("ADODB.Connection")
'����
objCon.Open(Application("ConnectionString"))
	'SQL�� ����
	strSql = "Select * From Upload Order By Num Desc"
	'���ڵ�� ��ü ����
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	'����¡ 2
	objRs.PageSize = 10 '//������ ������ ���� : ���������� 10��
	'���ڵ�� ��ü�� Select��ɾ� ���� : ����
	objRs.Open strSql, objCon, 1
		'��� ����� ���
		Call ShowRecordSet(objRs)
	'���ڵ�´ݱ�
	objRs.Close()
	'���ڵ�� ��ü ����
	Set objRs = Nothing
'�ݱ�
objCon.Close()
'Ŀ�ؼ� ��ü ����
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>
	<table border="1" width="100%">
		<tr><td>��ȣ</td><td>�� ��</td><td>�ۼ���</td><td>��ȸ��</td><td>����</td></tr>
<%
	If objRs.BOF Or objRs.EOF Then
%>
		<tr><td colspan="5" align="center">
		�Էµ� �����Ͱ� �����ϴ�.
		</td></tr>
<%
	Else
		'����¡ 3
		objRs.AbsolutePage = intPage '//���� ������ ������ ����
		'����¡ 4
		intPageCount = objRs.PageCount '//�� ������ �� ����
		'����¡ 5
		Dim i: i = 1 '//�ʱ��
		Do Until objRs.EOF Or i > objRs.PageSize '//���ǽ�
%>
		<tr><td><%=objRs("Num")%></td>
		<td width="350">
		<a href="BoardView.asp?Num=<%=objRs("Num")%>">
		<%=objRs("Title")%>
		</a>
		</td>
		<td><%=objRs("Name")%></td><td><%=objRs("ReadCount")%></td><td>
		<!--���� ���� �ٿ�ε� ����-->
		<a href="./BoardDown.asp?strFileName=<%=objRs("FileName")%>"><img src="./images/ext_zip.gif" border="0" 
	alt="<%=objRs("FileName")%>(<%=objRs("FileSize")%>)"></a>
	
		</td></tr>
<%
			objRs.MoveNext()
			i = i + 1 '//������
		Loop
%>
		<tr><td colspan="5" align="center">
<%
		'����¡ 6 : ��� ����¡ �Լ� ȣ��
		Call AdvancedPaging(intPage, intPageCount)
%>
		</td></tr>
<%
	End If
%>
	<tr><td colspan="5" align="right"><a href="BoardWrite.asp">�۾���</a></td></tr>
	</table>
<%
End Sub
%>
<%
Sub AdvancedPaging(PageNo, NumPage)'PageNo:����������, NumPage:��ü������
%>
<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
     <FONT style="font-size: 9pt;">
        <font color="#c0c0c0">[</font>
     <%     If PageNo > 10 Then %>
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%= ((PageNo - 1) \ 10) * 10 %>">��</a>
     <%     Else %>
        <font color="#c0c0c0">��</font>
     <%     End If %>               
        <font color="#c0c0c0">|</font>
     <%     For i = (((PageNo - 1) \ 10) * 10 + 1) To ((((PageNo - 1) \ 10) + 1) * 10)
        If i > NumPage Then
               Exit For
        End If
        If i = Int(PageNo) Then
     %>
        <b><%= i %></b> <font color="#c0c0c0">|</font> 
     <%     Else %>
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%= i %>"><%= i %></a> <font color="#c0c0c0">|</font> 
     <%     End If %>
     <%     Next %>
     <%     If CInt(i) < CInt(NumPage) Then %>
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">��</a>
     <%     Else %>
        <font color="#c0c0c0">��</font>
     <%     End If %>
        <font color="#c0c0c0">]</font>
     </FONT>
<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
<%
End Sub
%> 

