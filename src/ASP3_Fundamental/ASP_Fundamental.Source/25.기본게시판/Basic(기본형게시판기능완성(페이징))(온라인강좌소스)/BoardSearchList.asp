<%
Dim objCon: Dim objRs: Dim strSql
Dim strSearchField: strSearchField = Request("SearchField")'필드
Dim strSearchQuery: strSearchQuery = Request("SearchQuery")'검색어
Dim intPage: intPage = Request("Page")
If intPage = "" Then
	intPage = 1
End If
Dim intPageCount: intPageCount = 0
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")

If strSearchField = "All" Then '// 전체검색(Name+Title+Content)
	strSql = "Select * From Basic Where Name Like '%" & strSearchQuery & "%' Or Title Like '%" & strSearchQuery & "%' Or Content Like '%" & strSearchQuery & "%' Order By Num Desc"
Else '// 부분 검색
	strSql = "Select * From Basic Where " & strSearchField & " Like '%" & strSearchQuery & "%' Order By Num Desc"
End If

Set objRs = Server.CreateObject("ADODB.RecordSet")
objRs.PageSize = 10
objRs.Open strSql, objCon, 1
	Call ShowRecordSet(objRs) '// 출력(BoardList.asp와 동일)
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
		<TD align="center">번&nbsp;호</TD>
		<TD align="center">제&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;목</TD>
		<TD align="center">이&nbsp;름</TD>
		<TD align="center">작성일</TD>
		<TD align="center">조&nbsp;회</TD>
	</TR>
	<%
	If objRs.BOF Or objRs.EOF Then
	%>
	<TR><TD colspan="5" align="center">검색된 데이터가 없습니다.</TD></TR>
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
			<a href="./BoardList.asp">검색 종료</a>
		</TD>
	</TR>
	</TABLE>
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
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?SearchField=<%=Request("SearchField")%>&SearchQuery=<%=Request("SearchQuery")%>&Page=<%= ((PageNo - 1) \ 10) * 10 %>">◀</a>
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
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?SearchField=<%=Request("SearchField")%>&SearchQuery=<%=Request("SearchQuery")%>&Page=<%= i %>"><%= i %></a> <font color="#c0c0c0">|</font> 
     <%     End If %>
     <%     Next %>
     <%     If CInt(i) < CInt(NumPage) Then %>
        <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?SearchField=<%=Request("SearchField")%>&SearchQuery=<%=Request("SearchQuery")%>&Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">▶</a>
     <%     Else %>
        <font color="#c0c0c0">▷</font>
     <%     End If %>
        <font color="#c0c0c0">]</font>
     </FONT>
<!--이전 10개, 다음 10개 페이징 처리 종료-->
<%
End Sub
%> 
