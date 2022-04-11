<%
Dim objCon, strSql, objRs, Page, i, PageCount
strSql = "Select * From gBook Order By UID Desc"
'[1] Page ���� ���� : ���� ���° �������� �����ٰ���...
If Request("Page") = "" Then
    Page = 1
Else
    Page = CInt(Request("Page"))
End If
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=GuestBook;Initial Catalog=GuestBook;Data Source=172.16.25.99")
Set objRs = Server.CreateObject("ADODB.Recordset")
'[2] PageSize �Ӽ� : �� �������� ��� �����ٰ���... 3
objRs.PageSize = 3
'[3] 3��° �Ű������� 1 �Ǵ� 3���� �ʱ�ȭ
objRs.Open strSql, objCon, 1
'[!] ������ ī��Ʈ
PageCount = objRs.PageCount
If objRs.BOF Or objRs.EOF Then
    Response.Write("���Ͽ� �����Ͱ� �����ϴ�.")
Else
    '[4] AbsolutePage �Ӽ� : ���° ���������� �����ٰ���...
    objRs.AbsolutePage = Page
    '[5] i������ 1���� PageSize��ŭ�� ���
    i = 1
    Do Until objRs.EOF Or (i > objRs.PageSize) 
        Response.Write(objRs("UID") & objRs("User_Name") & "<br />")
        objRs.MoveNext()
        i = i + 1 
    Loop
    '[6] ����¡ �� : ����/����
    Call Paging(Page, PageCount)
    Call AdvancedPaging(Page, PageCount)
End If
objRs.Close()
Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>

<%
Sub Paging(Page, PageCount)
    If Page = 1 Then
%>
    <a href="#" style="color:Gray;">����</a>
<%  Else %>
    <a href="<%= Request.ServerVariables("SCRIPT_NAME") %>?Page=<%=Page - 1%>">����</a>
<%  End If %>
    (<%= Page %>/<%= PageCount %>)
<%  If Page = PageCount Then %>
    <a href="#" style="color:Gray;">����</a>
<%  Else %>
    <a href="<%= Request.ServerVariables("SCRIPT_NAME") %>?Page=<%=Page + 1 %>">����</a> 
<%  End If %>
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



