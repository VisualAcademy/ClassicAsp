<%
Dim objCon, strSql, objRs, Page, PageCount
strSql = "Select * From Basic Order By Num Desc"
'[1] 
If Request("Page") = "" Then
    Page = 1
Else
    Page = CInt(Request("Page"))
End If
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))
Set objRs = Server.CreateObject("ADODB.Recordset")
'[2] 
objRs.PageSize = 10
'[!] 1 �Ǵ� 3
objRs.Open strSql, objCon, 1
If objRs.BOF Or objRs.EOF Then ' �����Ͱ� ���ٸ�...
    Response.Write("�Էµ� �����Ͱ� �����ϴ�.")
Else
    '[3]    
    PageCount = objRs.PageCount
    '[4] 
    objRs.AbsolutePage = Page ' ���° ����������...
    Call ShowRecordset(objRs)
    Response.Write("<br />")
    Call AdvancedPaging(Page, PageCount)    
End If
objRs.Close()
Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>

<%
Sub AdvancedPaging(PageNo, NumPage)'PageNo:����������, NumPage:��ü������
%>
	<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
		<FONT style="font-size: 9pt;">
			<font color="#c0c0c0">[</font>
		<%	If PageNo > 10 Then %>
			<a href="./boardlist.asp?Page=<%= ((PageNo - 1) \ 10) * 10 %>">��</a>
		<%	Else %>
			<font color="#c0c0c0">��</font>
		<%	End If %>			
			<font color="#c0c0c0">|</font>
		<%	For i = (((PageNo - 1) \ 10) * 10 + 1) To ((((PageNo - 1) \ 10) + 1) * 10)
			If i > NumPage Then
				Exit For
			End If
			If i = Int(PageNo) Then
		%>
			<b><%= i %></b> <font color="#c0c0c0">|</font> 
		<%	Else %>
			<a href="./boardlist.asp?Page=<%= i %>"><%= i %></a> <font color="#c0c0c0">|</font> 
		<%	End If %>
		<%	Next %>
		<%	If CInt(i) < CInt(NumPage) Then %>
			<a href="./boardlist.asp?Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">��</a>
		<%	Else %>
			<font color="#c0c0c0">��</font>
		<%	End If %>
			<font color="#c0c0c0">]</font>
		</FONT>
	<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
<%
End Sub
%>


<%
Sub ShowRecordset(ByRef objRs)
%>
    <h3>�Խ��� ����Ʈ</h3>
    <table border="1" width="550">
        <tr>
            <td>����</td><td>�ۼ���</td>
        </tr>
<%  
    '[5] 
    Dim i: i = 1 
    Do Until objRs.EOF Or i > objRs.PageSize
%>
        <tr>
            <td>
            <a href="boardview.asp?Num=<%= objRs("Num") %>">
            <%= objRs.Fields("Title") %>
            </a>
            </td>
            <td><%= objRs("Name") %></td>
        </tr>
<% 
        objRs.MoveNext()
        i = i + 1
    Loop
%>
    </table>
<%
End Sub
%>

<br />
<a href="boardwrite.asp">�۾���</a>

<form action="boardsearchlist.asp" method="post">

<select name="SearchField">
    <option value="Name">�̸�</option>
    <option value="Title">����</option>
    <option value="Content">����</option>
</select>

<input type="text" name="SearchQuery" />

<input type="submit" value="�˻�" />

</form>















