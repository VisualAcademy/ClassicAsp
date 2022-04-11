<%
'--------------------------------------------------
' Title : Myado Education Version
' Program Name : ado_view.asp
' Program Description : 게시판 글의 내용 보여주기
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!--METADATA TYPE= "typelib" NAME= "ADODB Type Library" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll" -->

<%
Option Explicit

'변수 선언
Dim objDB, objRS, strSQL, strSQL2, ScrollAction, Num
Dim strDB

Num = Request("num")

'상위 목록으로 이동을 위한 변수
ScrollAction = Request.QueryString("scrollaction")

'--------------------------------------------------------------------
'[1] 레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[2] 데이터베이스 연결 문자열 생성
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=ZZAZIPKIDOTCOM"

'[3] 넘겨온 글번호에 해당되는 패스워드 가져오기
strSQL = "Select * From my_ado Where Num = " & Num
'Response.Write strSQL

'[4] 레코드셋 열기
objRs.Open strSQL, strDB, adOpenStatic, adLockPessimistic, adCmdText

With objRS
	
		.Fields("ReadCount") = .Fields("ReadCount") + 1   '조회수 1 증가

		.Update

		.Close

End With

Set objRS = Nothing
'--------------------------------------------------------------------

'--------------------------------------------------------------------
'레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'데이터베이스 처리(검색) 구문 작성
strSQL2 = "Select * From my_ado Where Num = " & Num

'레코드셋 열기
objRs.Open strSQL2, strDB, adOpenStatic, adLockPessimistic, adCmdText
'--------------------------------------------------------------------
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
						<a href="./ado_write_form.asp">[글쓰기]</a>&nbsp;
						<a href="./ado_modify.asp?num=<%=objRS.Fields("Num")%>">[수정]</a>&nbsp;
						<a href="./ado_delete.asp?num=<%=objRS.Fields("Num")%>">[삭제]</a>&nbsp;
						<a href="./ado_list.asp?intpage=<%=ScrollAction%>">[리스트]</a>
					</td>
				</tr>
				<% If Session("ID") = "Admin" Then %>
				<tr>
					<td colspan="2" align="center">
						<a href="./ado_delete_process.asp?num=<%=objRS.Fields("Num")%>&passwd=<%=objRS.Fields("Passwd")%>">현재글 지우기(<%=objRS.Fields("Passwd")%>)</a>
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

'레코드셋 개체의 인스턴스 해제
Set objRS = Nothing

%>
