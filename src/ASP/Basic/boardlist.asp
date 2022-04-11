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
'[!] 1 또는 3
objRs.Open strSql, objCon, 1
If objRs.BOF Or objRs.EOF Then ' 데이터가 없다면...
    Response.Write("입력된 데이터가 없습니다.")
Else
    '[3]    
    PageCount = objRs.PageCount
    '[4] 
    objRs.AbsolutePage = Page ' 몇번째 페이지부터...
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
Sub AdvancedPaging(PageNo, NumPage)'PageNo:현재페이지, NumPage:전체페이지
%>
	<!--이전 10개, 다음 10개 페이징 처리 시작-->
		<FONT style="font-size: 9pt;">
			<font color="#c0c0c0">[</font>
		<%	If PageNo > 10 Then %>
			<a href="./boardlist.asp?Page=<%= ((PageNo - 1) \ 10) * 10 %>">◀</a>
		<%	Else %>
			<font color="#c0c0c0">◁</font>
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
			<a href="./boardlist.asp?Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">▶</a>
		<%	Else %>
			<font color="#c0c0c0">▷</font>
		<%	End If %>
			<font color="#c0c0c0">]</font>
		</FONT>
	<!--이전 10개, 다음 10개 페이징 처리 종료-->
<%
End Sub
%>


<%
Sub ShowRecordset(ByRef objRs)
%>
    <h3>게시판 리스트</h3>
    <table border="1" width="550">
        <tr>
            <td>제목</td><td>작성자</td>
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
<a href="boardwrite.asp">글쓰기</a>

<form action="boardsearchlist.asp" method="post">

<select name="SearchField">
    <option value="Name">이름</option>
    <option value="Title">제목</option>
    <option value="Content">내용</option>
</select>

<input type="text" name="SearchQuery" />

<input type="submit" value="검색" />

</form>















