<%
'[1] ���� ����
Dim objCon: Dim objRs: Dim strSql
'[!] Paging
Dim intPage: intPage = Request("Page")	'//���� ������ ��
If intPage = "" Then
	intPage = 1
End If
Dim intPageCount: intPageCount = 0 '//�� ������ ��
'[2] Ŀ�ؼ� ��ü ����
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] Ŀ�ؼ� ��ü ���� : Ŀ�ؼ� ��Ʈ��
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
	strSql = "Select * From Basic Order By Num Desc"
	Set objRs = Server.CreateObject("ADODB.RecordSet")
		'[!] Paging
		objRs.PageSize = 10 '//������ ������ ����
		objRs.Open strSql, objCon, 1
			'[8] ��� ����� ���
			Call ShowRecordSet(objRs)	'//ShowRecordset() �Լ� ����
		'[9] ���ڵ� �� ��ü �ݱ�
		objRs.Close()
		'[10] ���ڵ� �� ��ü ����
	Set objRs = Nothing
'[12] Ŀ�ؼ� ��ü �ݱ�
objCon.Close()
'[13] Ŀ�ؼ� ��ü ����
Set objCon = Nothing
%>

<%
Sub ShowRecordSet(objRs)
%>
	<TABLE border="1" width="100%">
	<TR height="25">
		<TD align="center">��&nbsp;ȣ</TD>
		<TD align="center">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</TD>
		<TD align="center">��&nbsp;��</TD>
		<TD align="center">�ۼ���</TD>
		<TD align="center">��&nbsp;ȸ</TD>
	</TR>
	<%
	If objRs.BOF Or objRs.EOF Then
	%>
	<TR><TD colspan="5" align="center">�Էµ� �����Ͱ� �����ϴ�.</TD></TR>
	<%
	Else
		'[!] Paging
		intPageCount = objRs.PageCount '// �� ������ �� ����
		'[!] Paging
		objRs.AbsolutePage = intPage '// ���° ���������� �����ٰ���
		'[!] Paging
		Dim i: i = 1 '// ī��Ʈ�� ���� ����� ���ÿ� 1�� �ʱ�ȭ
		'[!] Paging
		Do Until objRs.EOF Or i > objRs.PageSize '//������ ���������
	%>
	<TR onmouseover="this.style.backgroundColor='yellow';" onmouseout="this.style.backgroundColor='';"
	 height="25">
		<TD align="center"><%=objRs("Num")%></TD>		
		<TD>
			<a href="./BoardView.asp?Num=<%=objRs("Num")%>">
			<%=objRs("Title")%>
			</a>
		</TD>
		<TD align="center">
			<a href="mailto:<%=objRs("Email")%>">
			<%=objRs("Name")%>
			</a>
		</TD>		
		<TD align="center">
			<%=Left(objRs("PostDate"),10)%>
		</TD>
		<TD align="center"><%=objRs("ReadCount")%></TD>
	</TR>
	<%
			objRs.MoveNext()
			'[!] Paging
			i = i + 1 '//������
		Loop
	End IF
	%>
	<TR height="25">
		<TD colspan="5" align="center">
		<%Call AdvancedPaging(intPage, intPageCount)%>
		</TD>
	</TR>
	<TR height="25">
		<TD colspan="5" align="center">
		<!-- #include file="./BoardSearch.asp" -->
		</TD>
	</TR>
	<TR height="25">
		<TD colspan="5" align="right">
			<a href="./BoardWrite.asp">�۾���</a>
		</TD>
	</TR>
	</TABLE>
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


















