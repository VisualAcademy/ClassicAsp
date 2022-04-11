<%
'[1] 변수 선언
Dim objCon: Dim objRs: Dim strSql
'[!] Paging
Dim intPage: intPage = Request("Page")	'//현재 페이지 수
If intPage = "" Then
	intPage = 1
End If
Dim intPageCount: intPageCount = 0 '//총 페이지 수
'[2] 커넥션 객체 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] 커넥션 객체 열기 : 커넥션 스트링
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
	strSql = "Select * From Basic Order By Num Desc"
	Set objRs = Server.CreateObject("ADODB.RecordSet")
		'[!] Paging
		objRs.PageSize = 10 '//페이지 사이즈 결정
		objRs.Open strSql, objCon, 1
			'[8] 모양 만들어 출력
			Call ShowRecordSet(objRs)	'//ShowRecordset() 함수 생성
		'[9] 레코드 셋 객체 닫기
		objRs.Close()
		'[10] 레코드 셋 객체 해제
	Set objRs = Nothing
'[12] 커넥션 객체 닫기
objCon.Close()
'[13] 커넥션 객체 해제
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
	<TR><TD colspan="5" align="center">입력된 데이터가 없습니다.</TD></TR>
	<%
	Else
		'[!] Paging
		intPageCount = objRs.PageCount '// 총 페이지 수 저장
		'[!] Paging
		objRs.AbsolutePage = intPage '// 몇번째 페이지부터 보여줄건지
		'[!] Paging
		Dim i: i = 1 '// 카운트용 변수 선언과 동시에 1로 초기화
		'[!] Paging
		Do Until objRs.EOF Or i > objRs.PageSize '//페이지 사이즈까지
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
			i = i + 1 '//증감식
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
			<a href="./BoardWrite.asp">글쓰기</a>
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


















