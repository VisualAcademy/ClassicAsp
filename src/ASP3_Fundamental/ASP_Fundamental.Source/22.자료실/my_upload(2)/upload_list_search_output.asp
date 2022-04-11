<%
'--------------------------------------------------
' Title : Myupload Education Version
' Program Name : upload_list_search_output.asp
' Program Description : 게시판 검색 결과 보여주기
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
Option Explicit
'---------------------------------------------------------------------
'변수 선언
Dim strSearchField, strSearchQuery, objDB, objRS, strSQL, intTotalPage, intPage, intCount
'---------------------------------------------------------------------
'기본 / 넘겨온 페이지값 변수에 저장. 없으면 첫번째 페이지(1)를 보여줌.
intPage = Request.QueryString("intpage")
IF intPage = "" Then
	intPage = 1
End IF

'---------------------------------------------------------------------
'넘겨져온 검색키워드 2개의 값을 변수에 저장.
strSearchField = Request("strsearchfield")
strSearchQuery = Request("strsearchquery")
'---------------------------------------------------------------------

'Connection개체의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")

'데이터베이스 열기
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")
         
'레코드 셋 개체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet")

'---------------------------------------------------------------------
'SQL Query 생성
strSQL = "Select * From my_upload Where " & strSearchField & " Like '%" & _
         strSearchQuery & "%' Order By Num Desc"
'---------------------------------------------------------------------

'레코드셋의 페이지 사이트 결정
objRS.PageSize = 5

'레코드 셋 열기
objRS.Open strSQL, objDB, 1
%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
<%
IF objRS.BOF OR objRS.EOF Then
'---------------------------------------------------------------------
%>
		<p align="center">
			<%=strSearchQuery%>
			로 검색을 수행한 결과, 검색된 데이터가 없습니다.
		</p>
		<%
'---------------------------------------------------------------------		
Else

'총 페이지수 계산(레코드의 수를 나타내는 속성인 pagecount사용
intTotalPage = objRS.PageCount 

'현재 리스트 페이지에서 보여줄 페이지의 절대페이지 지정
objRS.AbsolutePage = intPage
%>
		<div align="center">
			<table border="1">
				<tr>
					<td colspan="5" align="center">
						검 색 결 과
					</td>
				</tr>
				<tr>
					<td align="center">번호</td>
					<td align="center">제목</td>
					<td align="center">작성자</td>
					<td align="center">작성일</td>
					<td align="center">조회</td>
				</tr>
	<%
	intCount = 1   '1부터 보여줄 레코드 수만큼 루프 반복시 사용될 카운트변수
	
	'데이터베이스 끝까지 값을 출력 또는 보여줄 레코드 수가 될때까지 반복.
	Do Until objRS.EOF OR intCount > objRS.PageSize   
	%>
				<tr>
					<td align="center"><%=objRS.Fields("Num")%></td>
					<td align="center"><a href="./upload_view.asp?num=<%=objRS.Fields("Num")%>"><%=objRS.Fields("Title")%></a></td>
					<td align="center"><a href="mailto:<%=objRS.Fields("Email")%>"><%=objRS.Fields("Name")%></a></td>
					<td align="center"><%=objRS.Fields("PostDate")%></td>
					<td align="center"><%=objRS.Fields("ReadCount")%></td>
				</tr>
	<%
	objRS.MoveNext
	'카운터변수값 증가	
	intCount = intCount + 1
	loop
	%>
				<tr>
					<td align="center" colspan="5">
	<%

	'이전 페이지 이동 링크 작성
		IF CInt(intPage) <> 1 Then
	%>
						<a href="upload_list_search_output.asp?intpage=<%=(intPage-1)%>&strsearchfield=<%=strSearchField%>&strsearchquery=<%=strSearchQuery%>">[이전]</a>
	<%
		Else
	%>
						<a href="#" style="color:gray;">[이전]</a>
	<% 
		End IF
	%>
	<%
	'현재 페이지 및 전체 페이지 출력.
	
	Response.Write("[현재:" & intPage & "/")
	Response.Write("전체:" & intTotalPage & "]")	
	
	%>
	<%
	'다음 페이지 이동 링크 작성
		IF CInt(intPage) <> CInt(intTotalPage) Then
	%>
						<a href="upload_list_search_output.asp?intpage=<%=(intPage+1)%>&strsearchfield=<%=strSearchField%>&strsearchquery=<%=strSearchQuery%>">[다음]</a>
	<%
		Else
	%>
						<a href="#" style="color:gray;">[다음]</a>
	<% 
		End IF
				'---------------------------------------------------------------------
	%>
					</td>
				</tr>
				<tr>
					<td align="center" colspan="5">
						<a href="./upload_list.asp">[검색완료]</a>
					</td>
				</tr>
			</table>
		</div>
		<%
				'---------------------------------------------------------------------		
End IF
%>
	</BODY>
</HTML>
<%
'레코드셋 닫기
objRS.Close

'데이터베이스 닫기
objDB.Close

'레코드셋 개체의 인스턴스 해제
Set objRS = Nothing

'데이터베이스 객체의 인스턴스 해제
Set objDB = Nothing
%>
