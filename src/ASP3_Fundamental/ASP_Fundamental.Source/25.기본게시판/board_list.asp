<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_list.asp
' Program Description : Select문을 이용한 게시판 리스트 보여주기
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
Option Explicit

'변수 선언
Dim objDB, objRS, strSQL, intPage, intTotalPage, intCount

'기본 / 넘겨온 페이지값 변수에 저장. 없으면 첫번째 페이지(1)를 보여줌.
intPage = Request.QueryString("intpage")
If intPage = "" Then
	intPage = 1
End If

'데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")

'데이터베이스 열기
objDB.open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

'레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'데이터베이스 처리(검색) 구문 작성
strSQL = "Select * From my_board Order By Num Desc"

'페이지사이즈 결정
objRS.PageSize = 5

'레코드셋 열기(adOpenKeyset : 레코드셋이 고정, 레코드셋이 생성된후에 다른사용자가 추가한 자료를 볼 수 없다.)
'objRS.Open strSQL, objDB, adOpenKeyset
objRS.Open strSQL, objDB, 1
%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
<%
If objRS.BOF OR objRS.EOF Then
%>
		<p align="center">
			입력된 데이터가 없습니다.
		</p>
<%
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
						▣마이 게시판▣
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
					<td align="center"><a href="./board_view.asp?num=<%=objRS.Fields("Num")%>&scrollaction=<%=intPage%>"><%=objRS.Fields("Title")%></a></td>
					<td align="center"><a href="mailto:<%=objRS.Fields("Email")%>"><%=objRS.Fields("Name")%></a></td>
					<td align="center"><%=objRS.Fields("PostDate")%></td>
					<td align="center"><%=objRS.Fields("ReadCount")%></td>
				</tr>
	<%
	objRS.MoveNext
	'카운터변수값 증가	
	intCount = intCount + 1
	Loop
	%>
				<tr>
					<td align="center" colspan="5">
	<%
	'이전 페이지 이동 링크 작성
		If CInt(intPage) <> 1 Then
	%>
		<a href="board_list.asp?intpage=<%=(intPage-1)%>">[이전]</a>
	<%
		Else
	%>
		<a href="#" style="color:gray;">[이전]</a>
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
		<a href="board_list.asp?intpage=<%=(intPage+1)%>">[다음]</a>
	<%
		Else
	%>
						<a href="#" style="color:gray;">[다음]</a>
	<% 
		End If
	%>
					</td>
				</tr>
				<tr>
					<td align="center" colspan="5">
						<!-- -------------------------------- 검색폼 구현 ------------------------------------- -->
						<div align="center">
							<table border="0">
								<form name="board_list_search_form" method="post" action="./board_list_search_output.asp">
									<tr>
										<td>
											검색 : <select name="strsearchfield" size="1">
												<option value="name">
													작성자</option>
												<option value="title">
													제목</option>
												<option value="content">
													내용</option>
											</select><input type="text" name="strsearchquery" size="20"> <input type="submit" value="검 색">
										</td>
									</tr>
								</form>
							</table>
						</div>
						<!-- -------------------------------- 검색폼 구현 ------------------------------------- -->
					</td>
				</tr>
				<tr>
					<td colspan="5" align="right">
						<a href="./board_write_form.asp">[글쓰기]</a>

						<% If Session("ID") <> "Admin" Then %>
						<a href="./board_admin.asp">[관리자 로그인]</a>
						<% Else %>
						<a href="./board_admin_logout.asp">[관리자 로그아웃]</a>
						<% End If %>

					</td>
				</tr>
			</table>
		</div>
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
