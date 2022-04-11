<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boardlist.asp
' Program Description : 글 리스트 페이지
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
'[1] 변수 선언
'Option Explicit
Dim objCon: Dim objRs: Dim strSql
Dim Page: Page = 1'//
If Request("Page") <> "" Then
	Page = Request("Page")'넘겨져온 페이지 값이 담긴다.
End If
Dim PageCount'//총 페이지 수 저장
'[2] 인스턴스 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] 오픈
objCon.Open(Application("ConnectionString"))
'[4] 레코드셋 객체의 인스턴스 생성
Set objRs = Server.CreateObject("ADODB.RecordSet")
'[5] 명령어 실행
strSql = "Select * From Basic Order By Num Desc"
objRs.PageSize = 5'//페이지 사이즈 결정
objRs.Open strSql, objCon, 1'//1 또는 3으로 지정
'[6] 출력
If objRs.BOF Or objRs.EOF Then
	Response.Write("데이터가 없습니다.")
Else
	PageCount = objRs.PageCount'//총 페이지 값 저장
	objRs.AbsolutePage = Page'//
	Call ShowRecordSet(objRs)
End If
'[7] Close()
objRs.Close(): objCon.Close()
'[8] Nothing
Set objRs = Nothing: Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>

	<center><h3>기본 게시판 리스트</h3></center>

	<table border=1 width="100%">
	<tr>
		<TD>번호</TD>
		<TD>제목</TD>
		<TD>작성자</TD>
		<TD>작성일</TD>
		<TD>조회수</TD>
	</tr>
<%
	Dim i: i = 1'//페이지 사이즈 카운트. 초기식
	Do Until objRs.EOF Or i > objRs.PageSize'//조건식
%>
	<TR onmouseover="this.style.backgroundColor='yellow';" onmouseout="this.style.backgroundColor='';">
		<TD><%=objRs("Num")%></TD>
		<TD><a href="./boardview.asp?Num=<%=objRs("Num")%>"><%=objRs("Title")%></a></TD>
		<TD><%=objRs("Name")%></TD>
		<TD><%=objRs("PostDate")%></TD>
		<TD><%=objRs("ReadCount")%></TD>
	</TR>
<%
		objRs.MoveNext()
		i = i + 1'//증감식
	Loop
%>
	<table>
<%
End Sub
%>
<form action="" method="get">
페이지 이동 : <input type="text" name="Page" size=3>
<input type="submit" value="이동">
</form>
<form action="./boardsearchlist.asp" method="post">

검색어 : <input type="text" name="SearchQuery">

<select name="SearchField">
	<option value="Title">제목</option>
	<option value="Name">이름</option>
	<option value="Content">내용</option>
</select>

<input type="submit" value="검색">

</form>

<br>

<a href="./boardwrite.asp">글쓰기</a>



<hr>

<%
If Page > 1 Then
%>
	<a href="./boardlist.asp?Page=<%=Page-1%>">
	[이전] 
	</a>
<%
Else
%>
	<a href="#" style="color:silver;">
	[이전] 
	</a>
<%
End If
Response.Write("[" & Page & "/" & PageCount & "]")
If CInt(Page) < CInt(PageCount) Then
%>
	<a href="./boardlist.asp?Page=<%=Page+1%>">
	[다음]
	</a>
<%
Else
%>
	<a href="#" style="color:silver;">
	[다음]
	</a>
<%
End If
%>

<hr>

<%
Call AdvancedPaging(Page, PageCount)
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

