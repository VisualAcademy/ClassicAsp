<%
'--------------------------------------------------------------------
' Title : MyMemo Education Version
' Program Name : memo_list_page.asp
' Program Description : 페이징  및 검색을 처리한 리스트
' Include Files : None
' Last Update : 
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.zzazipki.com/
'--------------------------------------------------------------------
%>
<%
Option Explicit

'변수 선언
Dim objDB, objRS, strSQL, intPage, intTotalPage, intCount

'--------------------------------------------------------------------
'기본 / 넘겨온 페이지값 변수에 저장. 없으면 첫번째 페이지(1)를 보여줌.
intPage = Request.QueryString("intpage")
IF intPage = "" Then
	intPage = 1
End IF
'--------------------------------------------------------------------

'데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")

'데이터베이스 열기
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'데이터베이스 처리(검색) 구문 작성
strSQL = "Select * From my_memo Order By Num Desc"

'--------------------------------------------------------------------
'RecordSet 객체의 PageSize 속성 사용 : 한 페이지의 사이즈 결정
objRS.PageSize = 5
'--------------------------------------------------------------------

'레코드셋 열기(sql쿼리, db연결문자열, 커서타입, 락타입, 옵션)
'objRS.Open strSQL, objDB,adOpenKeyset
'페이징 처리시 레코드셋의 커서타입은 반드시 1 또는 3의 값을 넣어야함.
objRS.Open strSQL, objDB, 1

%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
IF objRS.Bof or objRS.Eof then
%>
		<P align="center">
			입력된 데이터가 없습니다.
		</P>
		<%
Else

'---------------------------------------------------------------------
'총 페이지수 계산(레코드의 수를 나타내는 속성인 pagecount사용)
intTotalPage = objRS.PageCount 

'현재 리스트 페이지에서 보여줄 페이지의 절대페이지 지정(특정 페이지로 이동)
objRS.AbsolutePage = intPage
'---------------------------------------------------------------------

%>
		<DIV align="center">
			<TABLE border="1">
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
	'---------------------------------------------------------------------
	intCount = 1   '1부터 보여줄 레코드 수만큼 루프 반복시 사용될 카운트변수
	
	'데이터베이스 끝까지 값을 출력 또는 보여줄 레코드 수가 될때까지 반복.
	Do Until objRS.EOF OR intCount > objRS.PageSize   
	'---------------------------------------------------------------------	
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
	'---------------------------------------------------------------------	
	'카운터변수값 증가	
	intCount = intCount + 1
	'---------------------------------------------------------------------		
	Loop
	%>
				<TR>
					<TD align="center" colspan="4">
						<%
	'---------------------------------------------------------------------			
	'이전 페이지 이동 링크 작성
		IF CInt(intPage) <> 1 Then
	%>
						<A href="memo_list_page.asp?intpage=<%=(intPage-1)%>">[이전]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[이전]</A>
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
		IF CInt(intPage) <> CInt(intTotalPage) then
	%>
						<A href="memo_list_page.asp?intpage=<%=(intPage+1)%>">[다음]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[다음]</A>
						<% 
		End IF
	'---------------------------------------------------------------------				
	%>
					</TD>
				</TR>
			</TABLE>
		</DIV>
		<%
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
