<%
'--------------------------------------------------
' Title : MyADO Education Version
' Program Name : ado_list.asp
' Program Description : Select문을 이용한 게시판 리스트 보여주기
' Include Files : adovbs.inc
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<!-- #include file="./scripts/adovbs.inc" -->
<%
'변수 선언
'Option Explicit
Dim objRS   '레코드셋 객체의 인스턴스
Dim strSQL   '데이터베이스 검색 구문 저장
Dim strDB   '데이터베이스 연결 문자열 저장
Dim intPage   '현재 페이지의 번호 저장
Dim intTotalPage   '전체 페이지의 수 저장
Dim intCount   '임시 카운터 변수

'기본 / 넘겨온 페이지값 변수에 저장. 없으면 첫번째 페이지(1)를 보여줌.
intPage = Request.QueryString("intpage")
If intPage = "" Then
	intPage = 1
End If

'--------------------------------------------------------------------
'[1] 레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[2] 데이터베이스 연결 문자열 생성
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=ZZAZIPKIDOTCOM"

'[3] 데이터베이스 처리(검색) 구문 작성
strSQL = "Select * From my_ado Order By Num Desc"

'[4] 페이지사이즈 결정
objRS.PageSize = 5

'[5] 레코드셋 열기
objRs.Open strSQL, strDB, adOpenStatic, adLockReadOnly, adCmdtext
'--------------------------------------------------------------------
%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
If objRS.BOF OR objRS.EOF Then
%>
		<P align="center">
			입력된 데이터가 없습니다.
		</P>
		<%
Else

'총 페이지수 계산(레코드의 수를 나타내는 속성인 pagecount사용
intTotalPage = objRS.PageCount 

'현재 리스트 페이지에서 보여줄 페이지의 절대페이지 지정
objRS.AbsolutePage = intPage

%>
		<DIV align="center">
			<TABLE border="1">
				<TR>
					<TD colspan="5" align="center" bgcolor="LightBlue">
						▣마이 ADO 게시판▣
					</TD>
				</TR>
				<TR>
					<TD align="center">
						번호
					</TD>
					<TD align="center">
						제목
					</TD>
					<TD align="center">
						작성자
					</TD>
					<TD align="center">
						작성일
					</TD>
					<TD align="center">
						조회
					</TD>
				</TR>
				<%
	intCount = 1   '1부터 보여줄 레코드 수만큼 루프 반복시 사용될 카운트변수
	
	'데이터베이스 끝까지 값을 출력 또는 보여줄 레코드 수가 될때까지 반복.
	Do Until objRS.EOF OR intCount > objRS.PageSize   
	%>
				<TR>
					<TD align="center">
						<%=objRS.Fields("Num")%>
					</TD>
					<TD align="center">
						<A href="./ado_view.asp?num=<%=objRS.Fields("Num")%>&scrollaction=<%=intPage%>">
							<%=objRS.Fields("Title")%>
						</A>
					</TD>
					<TD align="center">
						<A href="mailto:<%=objRS.Fields("Email")%>">
							<%=objRS.Fields("Name")%>
						</A>
					</TD>
					<TD align="center">
						<%=objRS.Fields("PostDate")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("ReadCount")%>
					</TD>
				</TR>
				<%
	objRS.MoveNext
	'카운터변수값 증가	
	intCount = intCount + 1
	Loop
	%>
				<TR>
					<TD align="center" colspan="5">
						<%
	'이전 페이지 이동 링크 작성
		If CInt(intPage) <> 1 Then
	%>
						<A href="ado_list.asp?intpage=<%=(intPage-1)%>">[이전]</A>
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
						<A href="ado_list.asp?intpage=<%=(intPage+1)%>">[다음]</A>
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
					<TD align="center" colspan="5">
						<!-- -------------------------------- 검색폼 구현 ------------------------------------- -->
						<DIV align="center">
							<TABLE border="0">
								<FORM name="ado_list_search_form" method="post" action="./ado_list_search_output.asp">
									<TR>
										<TD>
											검색 : <SELECT name="strsearchfield" size="1">
												<OPTION value="name">
													작성자</OPTION>
												<OPTION value="title">
													제목</OPTION>
												<OPTION value="content">
													내용</OPTION>
											</SELECT><INPUT type="text" name="strsearchquery" size="20"> <INPUT type="submit" value="검 색">
										</TD>
									</TR>
								</FORM>
							</TABLE>
						</DIV>
						<!-- -------------------------------- 검색폼 구현 ------------------------------------- -->
					</TD>
				</TR>
				<TR>
					<TD colspan="5" align="right">
						<A href="./ado_write_form.asp">[글쓰기]</A>
						<% If Session("ID") <> "Admin" Then %>
						<A href="./ado_admin.asp">[관리자 로그인]</A>
						<% Else %>
						<A href="./ado_admin_logout.asp">[관리자 로그아웃]</A>
						<% End If %>
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

'레코드셋 개체의 인스턴스 해제
Set objRS = Nothing
%>
