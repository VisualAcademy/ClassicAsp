<%
'--------------------------------------
' MyMemo Education Version
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------
%>
<%
Option Explicit
'---------------------------------------------------------------------
'변수 선언
Dim strSearchField, strSearchQuery, objDB, objRS, strSQL, intTotalPage, intPage, intCount
'---------------------------------------------------------------------
'기본 / 넘겨온 페이지값 변수에 저장. 없으면 첫번째 페이지(1)를 보여줌.
intPage = Request.QueryString("intpage")
If intPage = "" then
	intPage = 1
End If

'---------------------------------------------------------------------
'넘겨져온 2개의 값을 변수에 저장.
strSearchField = Request("strsearchfield")
strSearchQuery = Request("strsearchquery")
'---------------------------------------------------------------------

'Connection개체의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

         
'레코드 셋 개체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet")

'---------------------------------------------------------------------
'SQL Query 생성
strSQL = "Select * From my_memo Where " & strSearchField & " Like '%" & _
         strSearchQuery & "%' Order By Num Desc"
'---------------------------------------------------------------------

objRS.PageSize = 5

'레코드 셋 열기
objRS.Open strSQL, objDB, 1

%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
If objRS.BOF or objRS.EOF then
'---------------------------------------------------------------------
%>
		<P align="center">
			<font color="red"><%=strSearchQuery%></font>
			로 검색을 수행한 결과, 검색된 데이터가 없습니다.
		</P>
		<P align="center">
			<input type="button" value="뒤로가기" onclick="history.go(-1);">
		</P>
		<%
'---------------------------------------------------------------------		
Else

'총 페이지수 계산(레코드의 수를 나타내는 속성인 pagecount사용
intTotalPage = objRS.PageCount 

'현재 리스트 페이지에서 보여줄 페이지의 절대페이지 지정
objRS.AbsolutePage = intPage

%>
		<DIV align="center">
			<TABLE border="1" width="100%">
				<TR>
					<TD align="center" colspan="4">
						▣ 메모 검색하기▣
					</TD>
				</TR>
				<TR>
					<TD align="center">
						번 호
					</TD>
					<TD align="center">
						이 름
					</TD>
					<TD align="center">
						메&nbsp;&nbsp;&nbsp;모
					</TD>
					<TD align="center">
						날 짜
					</TD>
				</TR>
				<%
	intCount = 1   '1부터 보여줄 레코드 수만큼 루프 반복시 사용될 카운트변수
	
	'데이터베이스 끝까지 값을 출력 또는 보여줄 레코드 수가 될때까지 반복.
	Do Until objRS.EOF OR intCount > objRS.pagesize   
	%>
				<TR>
					<TD align="center">
						<%=objRS.Fields("Num")%>
					</TD>
					<TD align="center">
						<A href="mailto:<%=objRS.Fields("Email")%>">
							<%=objRS.Fields("Name")%>
						</A>
					</TD>
					<TD align="center">
						<%=objRS.Fields("Title")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("PostDate")%>
					</TD>
				</TR>
				<%
	objRS.MoveNext
	'카운터변수값 증가	
	intCount = intCount + 1
	Loop
	%>
				<TR>
					<TD align="center" colspan="4">
						<%

	'이전 페이지 이동 링크 작성
		If CInt(intPage) <> 1 then
	%>
						<A href="./memo_list_search_output.asp?intpage=<%=(intPage-1)%>&strsearchfield=<%=strSearchField%>&strsearchquery=<%=strSearchQuery%>">
							[이전]</A>
						<%
		else
	%>
						<A href="#" style="color:gray;">[이전]</A>
						<% 
		End If
	%>
						<%
	'현재 페이지 및 전체 페이지 출력.
	
	Response.Write("[현재:" & intPage & "/")
	Response.Write("전체:" & intTotalPage & "]")	
	
	%>
						<%
	'다음 페이지 이동 링크 작성
		If CInt(intPage) <> CInt(intTotalPage) then
	%>
						<A href="./memo_list_search_output.asp?intpage=<%=(intPage+1)%>&strsearchfield=<%=strSearchField%>&strsearchquery=<%=strSearchQuery%>">
							[다음]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[다음]</A>
						<% 
		End If
	%>
					</TD>
				</TR>
				<TR>
					<TD align="center" colspan="4">
						<A href="./memo_list.asp">[검색완료]</A>
					</TD>
				</TR>
			</TABLE>
		</DIV>
		<%
End If
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
