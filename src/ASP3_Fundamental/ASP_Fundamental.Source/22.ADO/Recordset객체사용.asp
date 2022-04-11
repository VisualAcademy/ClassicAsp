<%
'--------------------------------------------------------------------
' Title : ADO - Recordset객체
' Program Name : Recordset객체사용.asp
' Program Description : Recordset객체의 모든 멤버 사용
' Include Files : msado15.dll
' Copyright (C) 2002 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<!-- METADATA TYPE="typelib" FILE="C:\program files\common files\system\ado\msado15.dll"-->
<%
Dim objRs

'기본 / 넘겨온 페이지값 변수에 저장. 없으면 첫번째 페이지(1)를 보여줌.
intPage = Request.QueryString("intpage")
If intPage = "" Then
	intPage = 1
End If

'[1] 레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[2] 데이터베이스 연결 문자열 생성
strDB = "Provider=SQLOLEDB.1;Password=sample;Persist Security Info=True;User ID=sample;Initial Catalog=sample;Data Source=(local)"

'[3] 데이터베이스 처리(검색) 구문 작성
strSQL = "Select * From my_memo Order By Num Desc"

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

'[]현재 레코드의 페이지를 설정. 현재 리스트 페이지에서 보여줄 페이지의 절대페이지 지정
objRS.AbsolutePage = intPage

'[]현재 레코드의 위치를 설정.
Response.Write("<center>현재 레코드의 위치 : " & objRs.AbsolutePosition & "</center>")
Response.End
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
