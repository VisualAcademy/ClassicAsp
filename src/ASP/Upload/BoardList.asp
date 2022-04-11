<%
'변수 선언 부
Dim objCon: Dim objRs: Dim strSql
'페이징 1
Dim intPage: intPage = Request("Page")
If intPage = "" Then
	intPage = 1
End If
Dim intPageCount: intPageCount = 0
'커넥션 객체 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'연결
objCon.Open(Application("ConnectionString"))
	'SQL문 지정
	strSql = "Select * From Upload Order By Num Desc"
	'레코드셋 객체 생성
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	'페이징 2
	objRs.PageSize = 10 '//페이지 사이즈 지정 : 한페이지에 10개
	'레코드셋 객체로 Select명령어 실행 : 열기
	objRs.Open strSql, objCon, 1
		'모양 만들어 출력
		Call ShowRecordSet(objRs)
	'레코드셋닫기
	objRs.Close()
	'레코드셋 객체 해제
	Set objRs = Nothing
'닫기
objCon.Close()
'커넥션 객체 해제
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>
	<table border="1" width="100%">
		<tr><td>번호</td><td>제 목</td><td>작성자</td><td>조회수</td><td>파일</td></tr>
<%
	If objRs.BOF Or objRs.EOF Then
%>
		<tr><td colspan="5" align="center">
		입력된 데이터가 없습니다.
		</td></tr>
<%
	Else
		'페이징 3
		objRs.AbsolutePage = intPage '//현재 보여줄 페이지 지정
		'페이징 4
		intPageCount = objRs.PageCount '//총 페이지 수 저장
		'페이징 5
		Dim i: i = 1 '//초기식
		Do Until objRs.EOF Or i > objRs.PageSize '//조건식
%>
		<tr><td><%=objRs("Num")%></td>
		<td width="350">
		<a href="BoardView.asp?Num=<%=objRs("Num")%>">
		<%=objRs("Title")%>
		</a>
		</td>
		<td><%=objRs("Name")%></td><td><%=objRs("ReadCount")%></td><td>
		<!--파일 강제 다운로드 적용-->
		<a href="./BoardDown.asp?strFileName=<%=objRs("FileName")%>"><img src="./images/ext_zip.gif" border="0" 
	alt="<%=objRs("FileName")%>(<%=objRs("FileSize")%>)"></a>
	
		</td></tr>
<%
			objRs.MoveNext()
			i = i + 1 '//증감식
		Loop
%>
		<tr><td colspan="5" align="center">
<%
		'페이징 6 : 고급 페이징 함수 호출
		Call AdvancedPaging(intPage, intPageCount)
%>
		</td></tr>
<%
	End If
%>
	<tr><td colspan="5" align="right"><a href="BoardWrite.asp">글쓰기</a></td></tr>
	</table>
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

