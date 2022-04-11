<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_view.asp
' Program Description : 게시판 글의 내용 보여주기
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
Option Explicit

'변수 선언
Dim objDB, objRS, strSQL, strSQL2, ScrollAction, Num

Num = Request.QueryString("num")

'상위 목록으로 이동을 위한 변수
ScrollAction = Request.QueryString("scrollaction")

'데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")

'데이터베이스 열기
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

'해당 번호의 조회수 증가시키는 쿼리 작성
strSQL = "Update my_board Set ReadCount = ReadCount + 1 "
strSQL = strSQL & "Where Num = " & Num
'디버깅
'Response.Write strSQL

'쿼리 실행
objDB.Execute strSQL

'레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'데이터베이스 처리(검색) 구문 작성
strSQL2 = "Select * From my_board Where Num = " & Num

'레코드셋 열기(adOpenKeyset : 레코드셋이 고정, 레코드셋이 생성된후에 다른사용자가 추가한 자료를 볼 수 없다.)
'objRS.Open strSQL, objDB, adOpenKeyset
objRS.Open strSQL2, objDB, 1
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
%>
		<div align="center">
			<table border="1">
				<tr>
					<td colspan="2" align="center">
						▣마이 게시판 글보기▣
					</td>
				</tr>
				<tr>
					<td align="center">제목</td>
					<td align="center"><%=objRS.Fields("Title")%></td>
				</tr>
				<tr>
					<td align="center">작성자</td><td align="center"><%=objRS.Fields("Name")%></td>
				</tr>
				<tr>
					<td align="center">작성일</td><td><%=objRS.Fields("PostDate")%></td>
				</tr>
				<tr>
					<td align="center">조회</td><td><%=objRS.Fields("ReadCount")%></td>
				</tr>
				<tr>
					<td align="center" colspan="2"><%=Replace(objRS.Fields("Content"),Chr(13),"<br>")%></td>
				</tr>
				<tr>
					<td colspan="2" align="right">
						<a href="./board_write_form.asp">[글쓰기]</a>&nbsp;
						<a href="./board_modify.asp?num=<%=objRS.Fields("Num")%>">[수정]</a>&nbsp;
						<a href="./board_delete.asp?num=<%=objRS.Fields("Num")%>">[삭제]</a>&nbsp;
						<a href="./board_list.asp?intpage=<%=ScrollAction%>">[리스트]</a>
					</td>
				</tr>
				<% If Session("ID") = "Admin" Then %>
				<tr>
					<td colspan="2" align="center">
						<a href="./board_delete_process.asp?num=<%=objRS.Fields("Num")%>&passwd=<%=objRS.Fields("Passwd")%>">현재글 지우기(<%=objRS.Fields("Passwd")%>)</a>
					</td>
				</tr>
				<% End If %>
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
