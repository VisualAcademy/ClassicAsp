<%
Dim objCon, strSql, objRs, Page, i, PageCount
strSql = "Select * From gBook Order By UID Desc"
'[1] Page 변수 설정 : 현재 몇번째 페이지를 보여줄건지...
If Request("Page") = "" Then
    Page = 1
Else
    Page = CInt(Request("Page"))
End If
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=GuestBook;Initial Catalog=GuestBook;Data Source=172.16.25.99")
Set objRs = Server.CreateObject("ADODB.Recordset")
'[2] PageSize 속성 : 한 페이지에 몇개씩 보여줄건지... 3
objRs.PageSize = 1
'[3] 3번째 매개변수를 1 또는 3으로 초기화
objRs.Open strSql, objCon, 1
'[!] 페이지 카운트
PageCount = objRs.PageCount
If objRs.BOF Or objRs.EOF Then
    Response.Write("방명록에 데이터가 없습니다.")
Else
    '[4] AbsolutePage 속성 : 몇번째 페이지부터 보여줄건지...
    objRs.AbsolutePage = Page
%>
    <h3 align="center">방명록 리스트</h3>
    <table border="1" width="500" align="center">
<%
    '[5] i변수를 1부터 PageSize만큼만 출력
    i = 1
    Do Until objRs.EOF Or (i > objRs.PageSize)
%>
        <tr><td>
        <%= objRs("User_Name") %>님이 
        <%= objRs("RegDate") %>에 남긴 글입니다.<br />
        <%= Replace(objRs("Comment"), Chr(13), "<br />") %><br />
        IP주소 : <%= objRs("User_IP") %><br />
        <a href="Update.asp?ID=<%= objRs("UID") %>">수정</a>
        <a href="Remove.asp?ID=<%= objRs("UID") %>">삭제</a>            
        </td></tr>
<%
        objRs.MoveNext()
        i = i + 1
    Loop

%>
    </table>
<%
    Call Paging(Page, PageCount)
    Response.Write("<br />")
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
    <a href="#" style="color:Gray;">이전</a>
<%  Else %>
    <a href="<%= Request.ServerVariables("SCRIPT_NAME") %>?Page=<%=Page - 1%>">이전</a>
<%  End If %>
    (<%= Page %>/<%= PageCount %>)
<%  If Page = PageCount Then %>
    <a href="#" style="color:Gray;">다음</a>
<%  Else %>
    <a href="<%= Request.ServerVariables("SCRIPT_NAME") %>?Page=<%=Page + 1 %>">다음</a> 
<%  End If %>
<%
End Sub
%>

<%
Sub AdvancedPaging(PageNo, NumPage)'PageNo:현재페이지, NumPage:전체페이지
%>
<!--이전 10개, 다음 10개 페이징 처리 시작-->
	<FONT style="font-size: 9pt;">
	   <font color="#c0c0c0">[</font>
	<%     If PageNo > 10 Then %>
	   <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%= ((PageNo - 1) \ 10) * 10 %>">◀</a>
	<%     Else %>
	   <font color="#c0c0c0">◁</font>
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
	   <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">▶</a>
	<%     Else %>
	   <font color="#c0c0c0">▷</font>
	<%     End If %>
	   <font color="#c0c0c0">]</font>
	</FONT>
<!--이전 10개, 다음 10개 페이징 처리 종료-->
<%
End Sub
%>


