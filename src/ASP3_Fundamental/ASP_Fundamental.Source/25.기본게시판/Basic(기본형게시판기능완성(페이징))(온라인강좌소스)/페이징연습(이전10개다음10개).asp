<%
Dim objCon: Dim objRs: Dim strSql
'[!]����¡ ó�� ����
Dim intPage: intPage = Request("Page")
If intPage = "" Then	'/�Ѱ��� �� ������ ���� ������ 1���������� ���
	intPage = 1
End If
'[!]����/����
Dim intPageCount: intPageCount = 0 '//�� ������ ��
Set objCon = Server.CreateObject("ADODB.Connection")
	objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
		strSql = "Select * From Basic Order By Num Desc"
		Set objRs = Server.CreateObject("ADODB.RecordSet")
			'[!] �� �������� 5���� �����ֱ�
			objRs.PageSize = 5
			'[!] ����¡ ó���� �ݵ�� 3��° ����(Ŀ��Ÿ��) : 1 �Ǵ� 3
			objRs.Open strSql, objCon, 1
				'���
				If objRs.BOF Or objRs.EOF Then
				Else
					'[!]����/����
					intPageCount = objRs.PageCount
					'[!] �׷�, ������ �о�� �� �� ���°���� �����ٰ���
					objRs.AbsolutePage = intPage
					Dim i: i = 1	'//�ʱ��
					Do Until objRs.EOF Or i > objRs.PageSize '//���ǽ�
						Response.Write(objRs.Fields("Num") & "<br>")
						objRs.MoveNext()
						i = i + 1 '//������
					Loop
				End If
			objRs.Close()
		Set objRs = Nothing
	objCon.Close()
Set objCon = Nothing
%>

<hr>

<%
Call AdvancedPaging(intPage, intPageCount)'/��� ����¡ �Լ� ȣ��
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


<hr>

<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
������ : 
<input type="text" name="Page" value="<%=intPage%>" size="3" style="text-align:center;"> 
<input type="submit" value="�̵�">
</form>

<hr>

<%
If intPage > 1 Then
%>
	<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%=intPage-1%>">
	[����]
	</a>
<%
Else
%>
	<span style="color:gray;">[����]</span>
<%
End If
%>
(����:<%=intPage%>/��:<%=intPageCount%>) 
<%
If CInt(intPage) <> CInt(intPageCount) Then
%>
	<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%=intPage+1%>">
	[����]
	</a>
<%
Else
%>
	<span style="color:gray;">[����]</span>
<%
End If
%>