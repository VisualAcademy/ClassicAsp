<%
'--------------------------------------------------------------------
' Title : MyMemo Education Version
' Program Name : memo_list.asp
' Program Description : 메모글 리스트 페이지
' Include Files : None
' Last Update : 2001.08.26.
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<%
'[1] 변수의 명시적 선언
Option Explicit

'[2] 변수 선언
Dim objDB			'Connection객체의 인스턴스 객체를 저장할 변수
Dim objRS			'RecordSet객체의 인스턴스 객체를 저장할 변수 
Dim strSQL			'MY메모장 DB에서 데이터를 검색하는 SQL문을 저장하는 변수		
Dim intPage			'MY메모장 현재 페이지 번호를 저장할 변수
Dim intTotalPage	'MY메모장 전체 페이지 번호를 저장할 변수 
Dim intCount		'1부터 보여줄 레코드 수만큼 루프 반복시 사용될 카운트변수

'--------------------------------------------------------------------
'[3] 변수에 값 대입 : 만약 전송된 페이지값이 있으면 현재 페이지 변수에 저장하고, 
'그렇지 않으면 첫번째 페이지(1)를 보여줌.
intPage = Request.QueryString("intPage")
If intPage = "" Then
	intPage = 1
End if
'--------------------------------------------------------------------

'[4] 데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")

'[5] 데이터베이스 열기(ODBC설정의 DSN정보와 SQL Server의 로그인 정보(uid/pwd) 사용)
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'[6] 레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[7] 데이터베이스 처리(검색) 구문 작성
strSQL = "Select * From my_memo Order By Num DESC"

'--------------------------------------------------------------------
'[8] RecordSet 객체의 PageSize 속성 사용 : 한 페이지의 사이즈 결정
objRS.PageSize = 5
'--------------------------------------------------------------------

'[9] 레코드셋 열기
'레코드셋 열기(sql쿼리, db연결문자열, 커서타입, 락타입, 옵션)
'ADO상수 사용시(objRS.Open strSQL, objDB, adOpenStatic, adLockReadOnly, adCmdtext)
'페이징 처리시 레코드셋의 커서타입은 반드시 1 또는 3의 값을 넣어야함.
objRS.Open strSQL, objDB, 1

'[10] MY메모장 리스트 페이지 처리 부분 시작
%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
'[11] MY 메모장 DB에 데이터가 없을 경우 구문 작성
If objRS.BOF OR objRS.EOF Then
%>
		<P align="center">
			입력된 데이터가 없습니다.
		</P>
		<P align="center">
			<A href="./memo_write_form.asp">[글쓰기]</A>
		</P>
		<%
Else

'--------------------------------------------------------------------
'[12] 총 페이지수 계산(레코드의 수를 나타내는 속성인 PageCount사용)
intTotalPage = objRS.PageCount 

'[13] 현재 리스트 페이지에서 보여줄 페이지의 절대페이지 지정(특정 페이지로 이동)
objRS.AbsolutePage = intPage
'--------------------------------------------------------------------
%>
		<DIV align="center">
			<TABLE border="1" width="100%">
				<TR>
					<TD colspan="5" align="center">
						<nobr>
						▣MY 메모장▣
						</nobr>
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
'[14] MY 메모장 리스트 처리
	intCount = 1   '1부터 보여줄 레코드 수만큼 루프 반복시 사용될 카운트변수
	
	'데이터베이스 끝까지 값을 출력 또는 보여줄 레코드 수가 될때까지 반복.
	Do Until objRS.EOF OR intCount > objRS.PageSize   
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
					<TD colspan="5" align="right">
						<A href="./memo_write_form.asp">[글쓰기]</A>&nbsp;&nbsp; 
						<A href="./memo_delete.asp">[지우기]]</A>
					</TD>
				</TR>
				<TR>
					<TD align="center" colspan="4">
						<%
'--------------------------------------------------------------------
'[15] 페이지 처리
	'이전 페이지 이동 링크 작성
		If CInt(intPage) <> 1 Then
	%>
						<A href="memo_list.asp?intpage=<%=(intPage-1)%>">[이전]</A>
						<%
		Else
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
		If CInt(intPage) <> CInt(intTotalPage) Then
	%>
						<A href="memo_list.asp?intpage=<%=(intPage+1)%>">[다음]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[다음]</A>
						<% 
		End If
		'--------------------------------------------------------------------

'[16] 검색폼 구현
	%>
					</TD>
				</TR>
				<TR>
					<TD align="center" colspan="4">
						<!-- -------------------------------- 검색폼 구현 ------------------------------------- -->
						<DIV align="center">
							<TABLE border="0">
								<FORM name="memo_list_search_form" method="post" action="./memo_list_search_output.asp">
									<TR>
										<TD>
											검색 : <SELECT name="strsearchfield" size="1">
												<OPTION value="name">
													이름</OPTION>
												<OPTION value="title">
													메모</OPTION>
											</SELECT><INPUT type="text" name="strsearchquery" size="20"> <INPUT type="submit" value="검 색">
										</TD>
									</TR>
								</FORM>
							</TABLE>
						</DIV>
						<!-- -------------------------------- 검색폼 구현 ------------------------------------- -->
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
'[17] 레코드셋 닫기
objRS.Close

'[18] 데이터베이스 닫기
objDB.Close

'[19] 레코드셋 개체의 인스턴스 해제
Set objRS = Nothing

'[20] 데이터베이스 객체의 인스턴스 해제
Set objDB = Nothing
%>
