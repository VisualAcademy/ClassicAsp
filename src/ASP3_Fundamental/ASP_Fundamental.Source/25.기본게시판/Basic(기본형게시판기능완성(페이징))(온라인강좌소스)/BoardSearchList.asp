<%
Dim objCon: Dim objRs: Dim strSql
Dim strSearchField: strSearchField = Request("SearchField")'�ʵ�
Dim strSearchQuery: strSearchQuery = Request("SearchQuery")'�˻���
Dim intPage: intPage = Request("Page")
If intPage = "" Then
	intPage = 1
End If
Dim intPageCount: intPageCount = 0
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")

If strSearchField = "All" Then '// ��ü�˻�(Name+Title+Content)
	strSql = "Select * From Basic Where Name Like '%" & strSearchQuery & "%' Or Title Like '%" & strSearchQuery & "%' Or Content Like '%" & strSearchQuery & "%' Order By Num Desc"
Else '// �κ� �˻�
	strSql = "Select * From Basic Where " & strSearchField & " Like '%" & strSearchQuery & "%' Order By Num Desc"
End If

Set objRs = Server.CreateObject("ADODB.RecordSet")
objRs.PageSize = 10
objRs.Open strSql, objCon, 1
	Call ShowRecordSet(objRs) '// ���(BoardList.asp�� ����)
objRs.Close()
Set objRs = Nothing
objCon.Close()
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
	<TR><TD colspan="5" align="center">�˻��� �����Ͱ� �����ϴ�.</TD></TR>
	<%
	Else
		intPageCount = objRs.PageCount
		objRs.AbsolutePage = intPage
		Dim i: i = 1
		Do Until objRs.EOF Or i > objRs.PageSize
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
			i = i + 1
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
			<a href="./BoardList.asp">�˻� ����</a>
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
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?SearchField=<%=Request("SearchField")%>&SearchQuery=<%=Request("SearchQuery")%>&Page=<%= ((PageNo - 1) \ 10) * 10 %>">��</a>
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
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?SearchField=<%=Request("SearchField")%>&SearchQuery=<%=Request("SearchQuery")%>&Page=<%= i %>"><%= i %></a> <font color="#c0c0c0">|</font> 
     <%     End If %>
     <%     Next %>
     <%     If CInt(i) < CInt(NumPage) Then %>
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?SearchField=<%=Request("SearchField")%>&SearchQuery=<%=Request("SearchQuery")%>&Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">��</a>
     <%     Else %>
        <font color="#c0c0c0">��</font>
     <%     End If %>
        <font color="#c0c0c0">]</font>
     </FONT>
<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
<%
End Sub
%> 
